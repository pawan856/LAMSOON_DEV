Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_STF_STFMain
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMAT")

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

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(TR_WH_FR, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
            uiFun.load_dropdown(TR_WH_TO, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            TR_BY.Text = Session("usr_id")
            If TR_DATE.Text = "" Then
                TR_DATE.Text = Now().Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Stock Transfer Maintenance"
            lbl_ImageHd.Text = "Transfered Items"
            lbl_TR_CODE.Text = "Transfered Code:"
            lbl_TR_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_TR_DATE.Text = "Date:"
            lbl_TR_BY.Text = "Transfered By:"
            lbl_TR_BATCH_NO.Text = "Batch No.:"
            lbl_TR_REF_NO.Text = "Ref. No.:"
            lbl_TR_WH_FR.Text = "From (Warehouse):"
            lbl_TR_WH_TO.Text = "To (Warehouse):"
            lbl_TR_REM.Text = "Remarks"

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
                TR_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Value = "貨品轉移維護"
            lbl_ImageHd.Text = "貨品詳情"
            lbl_TR_CODE.Text = "轉移號碼:"
            lbl_TR_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_TR_DATE.Text = "日期:"
            lbl_TR_BY.Text = "轉移者:"
            lbl_TR_BATCH_NO.Text = "批號:"
            lbl_TR_REF_NO.Text = "文件編號:"
            lbl_TR_WH_FR.Text = "由 (倉庫):"
            lbl_TR_WH_TO.Text = "到 (倉庫):"
            lbl_TR_REM.Text = "備註:"

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
                TR_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        TR_WH_FR.CssClass = "REQUIRED"
        TR_WH_TO.CssClass = "REQUIRED"
        TR_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'TR_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("TR_CODE") = ""
            ViewState("trd_loc_fr") = Nothing
            ViewState("trd_loc_to") = Nothing

            Call BindGV()
        Else
            dt = ViewState("dt")

        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSTF()
            'ElseIf moduleAction = "SELECTTOWH" Then
            'changeToPallet()
        End If

        cm = New CommonMenu("STF", lheader.Value, TR_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "Y"
        cm.haveAttachments = "Y"
        cm.haveNotes = "Y"
        cm.haveTasks = "Y"
        cm.haveEmail = "Y"
        cm.haveHistory = "Y"

        cm.genCM(cmBar)

        selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")

        If TR_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf TR_STATUS.Text = "POSTED" Then
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


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "No.", "編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Code.", "物料號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物料名稱")
                Call cU.changeGVLabel(oGridViewRow, e, "Qty", "數量")
                Call cU.changeGVLabel(oGridViewRow, e, "From<br />Batch No.", "From<br />Batch No.")
                Call cU.changeGVLabel(oGridViewRow, e, "From Loc.", "由位置")
                Call cU.changeGVLabel(oGridViewRow, e, "From Pallet No.", "由貨板")
                Call cU.changeGVLabel(oGridViewRow, e, "To<br />Batch No.", "To<br />Batch No.")
                Call cU.changeGVLabel(oGridViewRow, e, "To Loc.", "到位置")
                Call cU.changeGVLabel(oGridViewRow, e, "To Pallet No.", "到貨板")
                Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim
                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim

                CType(e.Row.FindControl("dsp_trd_batch_no_fr"), Label).Text = DataBinder.Eval(e.Row.DataItem, "trd_batch_no_fr").ToString.Trim
                CType(e.Row.FindControl("trd_batch_no_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_batch_no_fr").ToString.Trim

                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim
                CType(e.Row.FindControl("trd_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "trd_qty").ToString.Trim)

                CType(e.Row.FindControl("trd_loc_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_loc_fr").ToString.Trim
                CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).Text = DataBinder.Eval(e.Row.DataItem, "trd_loc_fr").ToString.Trim

                CType(e.Row.FindControl("trd_loc_to"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim
                CType(e.Row.FindControl("dsp_trd_loc_to"), Label).Text = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                CType(e.Row.FindControl("trd_pallet_no_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_fr").ToString.Trim
                CType(e.Row.FindControl("dsp_trd_pallet_no_fr"), Label).Text = DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_fr").ToString.Trim

                CType(e.Row.FindControl("trd_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "trd_rem").ToString.Trim

                CType(e.Row.FindControl("trd_pallet_no_to"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_to").ToString.Trim
                'Dim pDropDown As DropDownList = CType(e.Row.FindControl("trd_pallet_no_to"), DropDownList)

                'uiFun.load_dropdown(pDropDown, "select iloc_pallet_no from wms_item_loc_bal where " & _
                '                                "imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND " & _
                '                                "STORER_CODE = '" & STORER_CODE.SelectedValue.Trim & "' AND " & _
                '                                "ITM_CODE = '" & DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim & "' AND " & _
                '                                "PACK_KEY = '" & DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim & "' AND " & _
                '                                "ILOC_LOC = '" & DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim & "' AND " & _
                '                                "ISNULL(ILOC_PALLET_NO, '') <> '' " & _
                '                                " order by 1", "iloc_pallet_no", "iloc_pallet_no", "000", "000")

                'pDropDown.SelectedValue = gU.decodeNull(DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_to").ToString.Trim, "000")

                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp_FR"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp(1,'" & gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).ClientID) & "', '" & gU.jsHTMLEncode(CType(e.Row.FindControl("trd_loc_fr"), HiddenField).ClientID) & "')")
                nImage.Attributes.Add("onclick", "ItemLocLookUp(document.myform." & STORER_CODE.ClientID & ".value, " & _
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_fr"), HiddenField).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_pallet_no_fr"), Label).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_pallet_no_fr"), HiddenField).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_qty"), TextBox).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_batch_no_fr"), Label).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_batch_no_fr"), HiddenField).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_pallet_no_to"), TextBox).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_batch_no_to"), DropDownList).ClientID) & "' " & _
                                      ")")
                '"document.myform." & CType(e.Row.FindControl("itm_code"), TextBox).ClientID & ".value, " & _
                '"document.myform." & CType(e.Row.FindControl("pack_key"), TextBox).ClientID & ".value, " & _


                Dim tImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp_TO"), Image)
                tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                tImage.Attributes.Add("onclick", "LocLookUp(2,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_to"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_to"), HiddenField).ClientID) & "')")


                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                uiFun.load_dropdown(CType(e.Row.FindControl("trd_batch_no_to"), DropDownList), "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by dc_date_code", "DC_DATE_CODE", "DC_DATE_CODE", , Session("gSelectLabel"))
                If Not CType(e.Row.FindControl("trd_batch_no_to"), DropDownList).Items.FindByValue(DataBinder.Eval(e.Row.DataItem, "trd_batch_no_to").ToString.Trim) Is Nothing Then
                    CType(e.Row.FindControl("trd_batch_no_to"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "trd_batch_no_to").ToString.Trim
                End If

                'CType(e.Row.FindControl("trd_batch_no_to"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "trd_batch_no_to").ToString.Trim
                'CType(e.Row.FindControl("trd_batch_no_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_batch_no_fr").ToString.Trim

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
            Dim seq_string As String = "select MAX(CAST(trd_SEQ AS int)) + 1 from WMS_STOCK_ISSUE_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and TR_CODE = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("trd_seq") = ViewState("n_cur_seq").ToString
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


        If TR_WH_FR.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_WH_FR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_WH_FR.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If TR_WH_TO.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_WH_TO.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_WH_TO.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If TR_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_DATE.Text & "不能空白!", Session("gLang"))
            End If

            Return False
        ElseIf Not gU.isValidDate(TR_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_TR_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_TR_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim frloc As String = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                Dim toloc As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                Dim frPallet As String = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
                Dim toPallet As String = CType(GridView1.Rows(i).FindControl("trd_pallet_no_to"), TextBox).Text

                Dim frDateCode As String = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                Dim toDateCode As String = CType(GridView1.Rows(i).FindControl("trd_batch_no_to"), DropDownList).SelectedValue

                If frloc <> "" And toloc <> "" Then
                    If frloc = toloc And frPallet = toPallet And frDateCode = toDateCode Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Transfer Location Cannot Be Same!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "轉移不能一樣!", Session("gLang"))
                        End If

                        reloadHiddenValue(GridView1)

                        Return False
                    End If
                ElseIf frloc = "" And toloc <> "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "From Location Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "由位置不能空白!", Session("gLang"))
                    End If

                    reloadHiddenValue(GridView1)

                    Return False
                ElseIf frloc <> "" And toloc = "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "To Location Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "到位置不能空白!", Session("gLang"))
                    End If

                    reloadHiddenValue(GridView1)

                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text <> "" Then
                    If uiFun.gvValidate(Me, dt, "trd_qty", "Qty", _
                                     CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text) = False Then Return False


                    If CInt(CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text) <= 0 Then
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
            Next
        End If

        Return True

    End Function

    Private Sub reloadHiddenValue(ByRef gv As GridView)
        For i = 0 To gv.Rows.Count - 1
            'CType(gv.Rows(i).FindControl("trd_batch_no_fr"), TextBox).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), TextBox).Text
            CType(gv.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
        Next
    End Sub

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("STF", gConn, transaction)
                    'nextNo = TR_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from wms_stock_transfer " & _
                                "where tr_code = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code= '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code= '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_stock_transfer (" & _
                        "tr_code, imp_code, storer_code, " & _
                        "tr_status, tr_date, tr_by, " & _
                        "tr_batch_no, tr_ref_no, " & _
                        "tr_wh_fr, tr_wh_to, tr_rem, " & _
                         "sys_cb, sys_cd, sys_lub, sys_lud)" & _
                        "values ( " & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TR_STATUS.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(TR_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_BY.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TR_BATCH_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_REF_NO.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TR_WH_FR.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TR_WH_TO.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TR_REM.Text.Trim)) & "," & _
                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "trd_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into wms_stock_transfer_d (" & _
                                                "tr_code, imp_code, storer_code, " & _
                                                "trd_seq, itm_code, pack_key, " & _
                                                "trd_batch_no_fr, trd_qty, trd_loc_fr, " & _
                                                "trd_loc_to, trd_rem, " & _
                                                "trd_pallet_no_fr, trd_pallet_no_to, trd_batch_no_to," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, ""))) & ", " & _
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
                    sql_string = "update wms_stock_transfer set " & _
                                    "tr_status = " & gU.convdbNVCData(gU.dbEncode(TR_STATUS.Text)) & ", " & _
                                    "tr_date = " & gU.convdbDate(gU.dbEncode(TR_DATE.Text.Trim)) & ", " & _
                                    "tr_by = " & gU.convdbNVCData(gU.dbEncode(TR_BY.Text.Trim)) & ", " & _
                                    "tr_batch_no = " & gU.convdbNVCData(gU.dbEncode(TR_BATCH_NO.Text.Trim)) & ", " & _
                                    "tr_ref_no = " & gU.convdbNVCData(gU.dbEncode(TR_REF_NO.Text.Trim)) & ", " & _
                                    "tr_wh_fr = " & gU.convdbNVCData(gU.dbEncode(TR_WH_FR.SelectedValue)) & ", " & _
                                    "tr_wh_to = " & gU.convdbNVCData(gU.dbEncode(TR_WH_TO.SelectedValue)) & ", " & _
                                    "tr_rem = " & gU.convdbNVCData(gU.dbEncode(TR_REM.Text.Trim)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and TR_code = '" & gU.dbEncode(TR_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "trd_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_stock_transfer_d (" & _
                                                "tr_code, imp_code, storer_code, " & _
                                                "trd_seq, itm_code, pack_key, " & _
                                                "trd_batch_no_fr, trd_qty, trd_loc_fr, " & _
                                                "trd_loc_to, trd_rem, " & _
                                                "trd_pallet_no_fr, trd_pallet_no_to, trd_batch_no_to," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(TR_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, ""))) & ", " & _
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_stock_transfer_d " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and tr_code = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' " & _
                                            "and trd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update wms_stock_transfer_d set " & _
                                                "trd_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, ""))) & ", " & _
                                                "itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                "pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                "trd_batch_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                "trd_batch_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, ""))) & ", " & _
                                                "trd_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " & _
                                                "trd_loc_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " & _
                                                "trd_loc_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " & _
                                                "trd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_rem").ToString.Trim, ""))) & ", " & _
                                                "trd_pallet_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                "trd_pallet_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "000"))) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and tr_code = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' " & _
                                            "and trd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"
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
                    TR_CODE.Text = nextNo
                    ViewState("TR_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    TR_CODE.ForeColor = Drawing.Color.Black
                    TR_CODE.Font.Size = 10
                    'TR_CODE.CssClass = ""
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
        If ViewState("TR_CODE") <> "" Then
            pk_code = ViewState("TR_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("TR_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        TR_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            TR_STATUS.Text = "NEW"
            TR_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select TO_CHAR(wms_stock_transfer.TR_DATE,'DD/MM/YYYY') as TR_DATE,  wms_stock_transfer.* from wms_stock_transfer " & _
                        "where wms_stock_transfer.tr_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and wms_stock_transfer.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and wms_stock_transfer.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                TR_CODE.Text = dt.Rows(0).Item("TR_code").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                TR_STATUS.Text = dt.Rows(0).Item("TR_STATUS").ToString
                TR_DATE.Text = dt.Rows(0).Item("TR_DATE").ToString
                TR_BY.Text = dt.Rows(0).Item("TR_BY").ToString
                TR_BATCH_NO.Text = dt.Rows(0).Item("TR_BATCH_NO").ToString
                TR_REF_NO.Text = dt.Rows(0).Item("TR_REF_NO").ToString
                TR_WH_FR.SelectedValue = dt.Rows(0).Item("TR_WH_FR").ToString
                TR_WH_TO.SelectedValue = dt.Rows(0).Item("TR_WH_TO").ToString
                TR_REM.Text = dt.Rows(0).Item("TR_REM").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If TR_CODE.Text <> "" Then
                    'TR_CODE.ReadOnly = True
                    'TR_CODE.BorderWidth = 0
                    TR_CODE.BackColor = Drawing.Color.Transparent
                End If

                If TR_BY.Text <> "" Then
                    TR_BY.ReadOnly = True
                    TR_BY.BorderWidth = 0
                    TR_BY.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT wms_stock_transfer_d.*, wms_item.itm_name, 'U' as mFlag, wms_stock_transfer_d.trd_seq as old_seq " & _
                    "from wms_stock_transfer_d, wms_item " & _
                    "where wms_stock_transfer_d.imp_code = wms_item.imp_code " & _
                    "and wms_stock_transfer_d.storer_code = wms_item.storer_code " & _
                    "and wms_stock_transfer_d.itm_code = wms_item.itm_code " & _
                    "and wms_stock_transfer_d.PACK_KEY = wms_item.PACK_KEY " & _
                    "and wms_stock_transfer_d.tr_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_stock_transfer_d.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_stock_transfer_d.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by to_number(wms_stock_transfer_d.trd_seq)"
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
        Dim cancelSql As String = "update wms_stock_transfer " & _
                         "set TR_status = 'CANCELLED', " & _
                         "sys_lub = '" & Session("usr_id") & "', " & _
                         "sys_lud = Getdate() " & _
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and tr_code = '" & gU.dbEncode(TR_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            TR_STATUS.Text = "CANCELLED"

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
                Call save("Y")




                For Each rows As DataRow In dt.Rows

                    qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  to_char(max(ILOC_EXPIRY_DATE),'" & DDFORMAT & "') as ILOC_EXPIRY_DATE, TO_CHAR(max(ILOC_MANU_DATE),'" & DDFORMAT & "') as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                                "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                                "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                                "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                                "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("trd_loc_fr").ToString.Trim) & _
                                "' AND ILOC_PALLET_NO = '" & gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) & "' "

                    qtyTbl = gDB.getDataTable(qtyString)



                    If qtyTbl.Rows.Count > 0 Then
                        Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")

                        EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                        MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

                        If gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) = "000" Then
                            If CInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, "0")) > CInt(locQty) Then
                                qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                               "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                               "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                               "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                               "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("trd_loc_fr").ToString.Trim) & _
                               "' AND ILOC_PALLET_NO = '' "

                                qtyTbl = gDB.getDataTable(qtyString)

                                locQty = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item(0).ToString, "0")
                            End If
                        End If

                        If CInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, "0")) > CInt(locQty) Then
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
                            st.ITM_CODE = gU.decodeNull(rows.Item("itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "STF"
                            st.IO_DOC_ID = TR_CODE.Text.Trim
                            st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, ""), "0")
                            st.IO_CBM = 0
                            st.IO_KG = 0
                            st.IO_EXPIRY_DATE = EXP_DATE
                            st.IO_MANU_DATE = MANU_DATE

                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, "")

                            st.IO_WH = TR_WH_FR.SelectedValue
                            st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")
                            st.UpdateStockTrans("OUT", gConn, transaction)
                            st.UpdateStockBalTrans("OUT", gConn, transaction)

                            st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, "")
                            st.IO_WH = TR_WH_TO.SelectedValue
                            st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "")

                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, "")

                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)
                        End If
                    Else
                        imString = "SELECT M.*, D.* " & _
                                    "from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
                                    "where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                                    "AND M.ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & "' " & _
                                    "AND M.PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & "' " & _
                                    "AND M.IMP_CODE = D.IMP_CODE(+) " & _
                                    "AND M.STORER_CODE = D.STORER_CODE(+) " & _
                                    "AND M.PACK_KEY = D.PACK_KEY(+) " & _
                                    "AND M.ITM_CODE = D.ITM_CODE(+)"

                        imTbl = gDB.getDataTable(imString)

                        Dim balQty As String

                        If imTbl.Rows.Count > 0 Then
                            balQty = imTbl.Rows(0).Item("ITM_BALANCE").ToString
                        Else
                            balQty = "0"
                        End If

                        If gU.decodeEmptyCInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, ""), "0") <= CInt(balQty) Then
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNull(rows.Item("itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "STF"
                            st.IO_DOC_ID = TR_CODE.Text.Trim
                            st.IO_QTY = gU.decodeEmptyCInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, ""), "0")
                            st.IO_CBM = 0
                            st.IO_KG = 0

                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, "")

                            st.IO_WH = TR_WH_FR.SelectedValue
                            st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")
                            st.UpdateStockTrans("OUT", gConn, transaction)
                            st.UpdateStockBalTrans("OUT", gConn, transaction)

                            st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, "")
                            st.IO_WH = TR_WH_TO.SelectedValue
                            st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "")

                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, "")

                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)
                        Else
                            If transaction IsNot Nothing Then
                                transaction.Rollback()
                            End If

                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Transfer Qty cannot be Greater than Balance Qty!! \r\nCurrent Balance Qty:" & balQty, Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "轉移數量不能大於結餘數量\r\n現時結餘數量: " & balQty, Session("gLang"))
                            End If

                            Exit Sub
                        End If

                    End If
                Next

                updtSql = "update wms_stock_transfer " & _
                            "set TR_status = 'POSTED', " & _
                            "sys_lub = '" & Session("usr_id") & "', " & _
                            "sys_lud = Getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and tr_code = '" & gU.dbEncode(TR_CODE.Text) & "' "

                gDB.amendData(updtSql)

                transaction.Commit()

                TR_STATUS.Text = "POSTED"
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

    Protected Sub addItemtoSTF()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(TRD_SEQ AS int)) + 1 from wms_stock_transfer_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and TR_CODE = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' "
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
            SQLString += "and wms_item.itm_code || '_000_' || wms_item.pack_key || '_000_' || ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

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
                dt.Rows(rows_count - 1).Item("TRD_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("ITM_CODE") = stf_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("ITM_NAME") = stf_dt.Rows(i).Item("ITM_NAME")
                dt.Rows(rows_count - 1).Item("PACK_KEY") = stf_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("TRD_QTY") = stf_dt.Rows(i).Item("ITM_BALANCE")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next

            STORER_CODE.Enabled = False

            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
            Next

        End If
    End Sub

    Protected Sub TR_WH_FR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TR_WH_FR.SelectedIndexChanged
        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value = ""
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = ""

                'CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), TextBox).Text = ""

                CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value = ""
                CType(GridView1.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = ""
                CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value = ""
            Next
        End If
    End Sub

    Protected Sub TR_WH_TO_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TR_WH_TO.SelectedIndexChanged
        If GridView1.Rows.Count > 0 Then
            Dim toArray As New ArrayList
            For i = 0 To GridView1.Rows.Count - 1
                toArray.Add(CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value)

                CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value = ""
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = ""

                ViewState("trd_loc_tor") = toArray

                CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value


                CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value

                CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value


                CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value

                'Dim pDropDown As DropDownList = CType(GridView1.Rows(i).FindControl("trd_pallet_no_to"), DropDownList)

                'uiFun.load_dropdown(pDropDown, "select iloc_pallet_no from wms_item_loc_bal where " & _
                '                                "imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND " & _
                '                                "STORER_CODE = '" & STORER_CODE.SelectedValue.Trim & "' AND " & _
                '                                "ITM_CODE = '" & CType(GridView1.Rows(i).FindControl("itm_code"), TextBox).Text & "' AND " & _
                '                                "PACK_KEY = '" & CType(GridView1.Rows(i).FindControl("pack_key"), TextBox).Text & "' AND " & _
                '                                "ILOC_LOC = '" & CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value & "' AND " & _
                '                                "ISNULL(ILOC_PALLET_NO, '') <> '' " & _
                '                                " order by 1", "iloc_pallet_no", "iloc_pallet_no", "000", "000")
            Next
        End If
    End Sub

    'Protected Sub changeToPallet()
    '    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
    '        Dim currentIndex As Integer = CInt(selectedrowIndex.Value)

    '        'Dim pDropDown As DropDownList = CType(GridView1.Rows(currentIndex).FindControl("trd_pallet_no_to"), DropDownList)

    '        'uiFun.load_dropdown(pDropDown, "select iloc_pallet_no from wms_item_loc_bal where " & _
    '        '                                "imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND " & _
    '        '                                "STORER_CODE = '" & STORER_CODE.SelectedValue.Trim & "' AND " & _
    '        '                                "ITM_CODE = '" & CType(GridView1.Rows(currentIndex).FindControl("itm_code"), TextBox).Text & "' AND " & _
    '        '                                "PACK_KEY = '" & CType(GridView1.Rows(currentIndex).FindControl("pack_key"), TextBox).Text & "' AND " & _
    '        '                                "ILOC_LOC = '" & CType(GridView1.Rows(currentIndex).FindControl("trd_loc_to"), HiddenField).Value & "' AND " & _
    '        '                                "ISNULL(ILOC_PALLET_NO, '') <> '' " & _
    '        '                                " order by 1", "iloc_pallet_no", "iloc_pallet_no", "000", "000")

    '        CType(GridView1.Rows(currentIndex).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(currentIndex).FindControl("trd_loc_fr"), HiddenField).Value
    '        CType(GridView1.Rows(currentIndex).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(currentIndex).FindControl("trd_pallet_no_fr"), HiddenField).Value
    '        CType(GridView1.Rows(currentIndex).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(currentIndex).FindControl("trd_loc_to"), HiddenField).Value
    '    End If
    'End Sub
End Class
