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
            ViewState("pagemode") = Nothing
            ViewState("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))

            Select Case Session("PAGE_SESSION_MENU_CODE")
                Case "OP_RD"
                    uiFun.load_dropdown(TR_WH_FR, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
                    uiFun.load_dropdown(TR_WH_TO, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))

                    Dim tImage As Image = Image_Loc_LookUp_TR_TO_LOC
                    tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                    tImage.Attributes.Add("onclick", "LocLookUp(2,'','" & HttpUtility.HtmlEncode(dsp_TR_TO_LOC.ClientID) & "', '" & HttpUtility.HtmlEncode(TR_TO_LOC.ClientID) & "')")
                Case "OP_SRL", "OP_STF"
                    uiFun.load_dropdown(TR_WH_FR, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))
                    uiFun.load_dropdown(TR_WH_TO, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))

                    Dim tImage As Image = Image_Loc_LookUp_TR_TO_LOC
                    tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                    tImage.Attributes.Add("onclick", "LocLookUp(2,'','" & HttpUtility.HtmlEncode(dsp_TR_TO_LOC.ClientID) & "', '" & HttpUtility.HtmlEncode(TR_TO_LOC.ClientID) & "','T')")
                Case Else
                    uiFun.load_dropdown(TR_WH_FR, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))
                    uiFun.load_dropdown(TR_WH_TO, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))

                    Dim tImage As Image = Image_Loc_LookUp_TR_TO_LOC
                    tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                    tImage.Attributes.Add("onclick", "LocLookUp(2,'','" & HttpUtility.HtmlEncode(dsp_TR_TO_LOC.ClientID) & "', '" & HttpUtility.HtmlEncode(TR_TO_LOC.ClientID) & "')")
            End Select

        End If

        If ViewState("pagemode") = "N" Then
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


        Select Case Session("PAGE_SESSION_MENU_CODE")
            Case "OP_SRL", "OP_RD"
                selectItemBtn.Visible = True
                selectSTQ.Visible = False
                TQ_TR.Visible = False
            Case Else
                If ViewState("pagemode") = "N" Then
                    selectSTQ.Visible = True
                End If
                TR_TO_DRUM.Visible = False
                lbl_TR_TO_DRUM.Visible = False
                selectItemBtn.Visible = False
                TQ_TR.Visible = True
        End Select

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" Then
                lheader.Text = "Stock Relocation Maintenance"
            ElseIf Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
                lheader.Text = "Cable Re-Drum Maintenance"
            Else
                lheader.Text = "Inter-store Transfer Maintenance"
            End If

            lbl_ImageHd.Text = "Transfered Items"
            lbl_TR_CODE.Text = "Transfered Code:"
            lbl_TR_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_TR_EDI_RFT_NO.Text = "RFT No."
            lbl_TR_DATE.Text = "Date:"
            lbl_TR_BY.Text = "Transfered By:"
            lbl_TR_BATCH_NO.Text = "Lot:"
            lbl_TR_REF_NO.Text = "Ref. No.:"
            lbl_TR_WH_FR.Text = "From (Subinventory):"
            lbl_TR_WH_TO.Text = "To (Subinventory):"
            lbl_TR_TO_LOC.Text = "To Location"
            lbl_TR_TO_DRUM.Text = "To Drum"
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
            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting"")){getLoad();}else{return false;}"
            btnIssue.OnClientClick = "if (confirm(""Stock Balance will be updated.\r\n(Please confirm to Issue."")){getLoad();}else{return false;}"
            btnRcv.OnClientClick = "if (confirm(""Stock Balance will be updated?\r\n(Please confirm to Receive."")){getLoad();}else{return false;}"

            If ViewState("pagemode") = "N" Then
                TR_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨品轉移維護"
            lbl_ImageHd.Text = "貨品詳情"
            lbl_TR_CODE.Text = "轉移號碼:"
            lbl_TR_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_TR_EDI_RFT_NO.Text = "RFT No."
            lbl_TR_DATE.Text = "日期:"
            lbl_TR_BY.Text = "轉移者:"
            lbl_TR_BATCH_NO.Text = "批號:"
            lbl_TR_REF_NO.Text = "文件編號:"
            lbl_TR_WH_FR.Text = "由 (子庫存):"
            lbl_TR_WH_TO.Text = "至 (子庫存):"
            lbl_TR_TO_LOC.Text = "To"
            lbl_TR_TO_DRUM.Text = "To Drum"
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
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
            btnIssue.OnClientClick = "if (confirm(""Stock Balance will be updated.\r\n(Please confirm to Issue."")){getLoad();}else{return false;}"
            btnRcv.OnClientClick = "if (confirm(""Stock Balance will be updated?\r\n(Please confirm to Receive."")){getLoad();}else{return false;}"
            If ViewState("pagemode") = "N" Then
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

        If ViewState("pagemode") = "N" Then
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
            dsp_TR_TO_LOC.Text = TR_TO_LOC.Value

        ElseIf moduleAction = "SELECTTQ" Then

            addItemFromTQ()
            dsp_TR_TO_LOC.Text = TR_TO_LOC.Value

            'ElseIf moduleAction = "SELECTTOWH" Then
            'changeToPallet()
        End If

        cm = New CommonMenu("STF", lheader.Text, TR_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "Y"
        cm.haveAttachments = "Y"
        cm.haveNotes = "Y"
        cm.haveTasks = "Y"
        cm.haveEmail = "Y"
        cm.haveHistory = "Y"

        'cm.genCM(cmBar)

        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value, document.getElementById('" & TR_WH_FR.ClientID & "').value);")
        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OP_STF','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("TR_CODE") & "','N');")

        setPageCtrlAccess()

        If TR_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf TR_STATUS.Text = "POSTED" OrElse TR_STATUS.Text = "ISSUED" OrElse TR_STATUS.Text = "RECEIVED" Then
            ar.sec_viewMode = "Y"
            If TR_STATUS.Text = "ISSUED" Then
                exceptionEditList.Add("Image_Loc_LookUp_To")
                exceptionEditList.Add("saveBtn2")
                exceptionEditList.Add("saveBtn1")
                exceptionEditList.Add("btnRcv")
            End If
            btnPost.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

        selectSTQ.Attributes.Add("onclick", "TQLookUp();return false;")

        Select Case TR_STATUS.Text
            Case "NEW"
                btnIssue.Enabled = True
                btnRcv.Enabled = False
            Case "ISSUED"
                btnIssue.Enabled = False
                btnRcv.Enabled = True
            Case "RECEIVED"
                btnIssue.Enabled = False
                btnRcv.Enabled = False

            Case "CANCELLED"
                btnIssue.Enabled = False
                btnRcv.Enabled = False
            Case "POSTED"
                btnGenWO.Visible = False
                btnIssue.Visible = False
                btnRcv.Visible = False
        End Select

        Select Case Session("PAGE_SESSION_MENU_CODE")
            Case "OP_SRL", "OP_RD"
                If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then
                    For i = 0 To GridView1.Rows.Count - 1
                        GridView1.Columns(7).Visible = False
                    Next
                End If
            Case Else

        End Select
    End Sub

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

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

                CType(e.Row.FindControl("trd_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "trd_seq").ToString.Trim

                CType(e.Row.FindControl("itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim

                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim
                CType(e.Row.FindControl("TRD_UOM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TRD_UOM").ToString.Trim

                Dim pDropDown As DropDownList = CType(e.Row.FindControl("trd_batch_no_fr"), DropDownList)

                Dim ddSQL As String = ""
                Dim addSQL As String = ""
                Dim FromLoc As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_fr").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "trd_loc_fr").ToString.Trim <> "" Then addSQL &= "ILOC_LOC = '" & Right((FromLoc.ToString.Trim), 8) & "' AND ILOC_WH = '" & TR_WH_FR.SelectedValue.ToString.Trim & "' AND "

                ddSQL = "select DISTINCT iloc_batch_no from wms_item_loc_bal where " &
                                                "imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND " &
                                                "STORER_CODE = '" & STORER_CODE.SelectedValue.Trim & "' AND " &
                                                "ITM_CODE = '" & DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim & "' AND " &
                                                "PACK_KEY = '" & DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim & "' AND " &
                                                addSQL &
                                                "ISNULL(ILOC_BATCH_NO, '') <> '' " &
                                                " order by ILOC_BATCH_NO"



                uiFun.load_dropdown(pDropDown, ddSQL, "iloc_batch_no", "iloc_batch_no", , , DataBinder.Eval(e.Row.DataItem, "trd_batch_no_fr").ToString.Trim)

                FromLoc = TR_WH_FR.SelectedValue.ToString() + Right(FromLoc, 8)
                'If FromLoc.Length > 12 Then
                '    FromLoc = FromLoc.Substring(TR_WH_FR.SelectedValue.ToString().Length)
                'End If

                CType(e.Row.FindControl("trd_loc_fr"), HiddenField).Value = FromLoc
                CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).Text = FromLoc

                Dim ToLoc As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                ToLoc = TR_WH_TO.SelectedValue.ToString() + Right(ToLoc, 8)

                'If ToLoc.Length > 12 Then
                '    ToLoc = ToLoc.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                'End If

                ''for only QCFG AND FG01 WAREHOUSE'
                'If TR_WH_FR.SelectedValue = "QCFG" AndAlso TR_WH_TO.SelectedValue = "FG01" Then
                '    Dim ToLoc1 As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                '    ToLoc1 = TR_WH_TO.SelectedValue.ToString() + ToLoc

                '    If ToLoc1.Length > 12 Then
                '        ToLoc1 = ToLoc1.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                '        ToLoc = ToLoc1.Remove(ToLoc1.Length - 2, 2) + "00"
                '    End If
                'End If


                'for only QCFG,FGTM,FG01,FGLOT AND FGHD WAREHOUSE'
                If TR_WH_TO.SelectedValue = "QCFG" Then
                    Dim ToLocFG As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                    ToLocFG = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocFG.Length >= 12 Then
                        ToLocFG = ToLocFG.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocFG.Remove(ToLocFG.Length - 2, 2) + "FG"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGTM" Then
                    Dim ToLocTM As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "TM"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FG01" Then
                    Dim ToLocTM As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "00"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGLOT" Then
                    Dim ToLocTM As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "LT"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGHD" Then
                    Dim ToLocHD As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                    ToLocHD = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocHD.Length >= 12 Then
                        ToLocHD = ToLocHD.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocHD.Remove(ToLocHD.Length - 2, 2) + "HD"
                    End If
                End If

                CType(e.Row.FindControl("trd_loc_to"), HiddenField).Value = ToLoc
                CType(e.Row.FindControl("dsp_trd_loc_to"), Label).Text = ToLoc

                CType(e.Row.FindControl("TRD_ITM_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TRD_ITM_NAME").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("trd_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "trd_qty").ToString.Trim)
                CType(e.Row.FindControl("trd_qty2"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "trd_qty2").ToString.Trim)

                CType(e.Row.FindControl("trd_pallet_no_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_fr").ToString.Trim
                CType(e.Row.FindControl("dsp_trd_pallet_no_fr"), Label).Text = DataBinder.Eval(e.Row.DataItem, "trd_pallet_no_fr").ToString.Trim

                CType(e.Row.FindControl("trd_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "trd_rem").ToString.Trim

                CType(e.Row.FindControl("TRD_EXPIRY_DATE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TRD_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("TRD_MANU_DATE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TRD_MANU_DATE").ToString.Trim

                CType(e.Row.FindControl("TRD_SERIAL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_SERIAL").ToString.Trim
                CType(e.Row.FindControl("TRD_SERIAL"), TextBox).Attributes.Add("readonly", "readonly")

                CType(e.Row.FindControl("TRD_UOM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TRD_UOM").ToString.Trim
                CType(e.Row.FindControl("TRD_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TRD_UOM2").ToString.Trim

                CType(e.Row.FindControl("TRD_ORG_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_ORG_QTY").ToString.Trim
                CType(e.Row.FindControl("TRD_ORG_QTY"), TextBox).Attributes.Add("readonly", "readonly")
                CType(e.Row.FindControl("TRD_ORG_QTY2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_ORG_QTY2").ToString.Trim
                CType(e.Row.FindControl("TRD_ORG_QTY2"), TextBox).Attributes.Add("readonly", "readonly")

                CType(e.Row.FindControl("TRD_REQ_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_REQ_QTY").ToString.Trim
                CType(e.Row.FindControl("TRD_REQ_QTY"), TextBox).Attributes.Add("readonly", "readonly")

                CType(e.Row.FindControl("TRD_DRUM_ID_FR"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_DRUM_ID_FR").ToString.Trim
                CType(e.Row.FindControl("TRD_DRUM_ID_FR"), TextBox).Attributes.Add("readonly", "readonly")
                CType(e.Row.FindControl("TRD_DRUM_LV_FR"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_DRUM_LV_FR").ToString.Trim
                CType(e.Row.FindControl("TRD_DRUM_LV_FR"), TextBox).Attributes.Add("readonly", "readonly")

                CType(e.Row.FindControl("TRD_DRUM_ID_TO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_DRUM_ID_TO").ToString.Trim
                CType(e.Row.FindControl("TRD_DRUM_LV_TO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TRD_DRUM_LV_TO").ToString.Trim



                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp_FR"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")


                Dim tImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp_TO"), Image)
                Select Case Session("PAGE_SESSION_MENU_CODE")
                    Case "OP_RD"
                        nImage.Attributes.Add("onclick", "LocLookUp(1,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_fr"), HiddenField).ClientID) & "','')")
                        tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                        tImage.Attributes.Add("onclick", "LocLookUp(2,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_to"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_to"), HiddenField).ClientID) & "','')")

                    Case "OP_SRL"
                        nImage.Attributes.Add("onclick", "LocLookUp(1,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_fr"), HiddenField).ClientID) & "','F')")
                        tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                        tImage.Attributes.Add("onclick", "LocLookUp(2,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_to"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_to"), HiddenField).ClientID) & "','T')")

                    Case Else
                        nImage.Attributes.Add("onclick", "ItemLocLookUp(document.myform." & STORER_CODE.ClientID & ".value, " &
                                              "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim) & "', " &
                                              "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_fr"), Label).ClientID) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_fr"), HiddenField).ClientID) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_pallet_no_fr"), Label).ClientID) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_pallet_no_fr"), HiddenField).ClientID) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_org_qty"), TextBox).ClientID) & "', " &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_batch_no_fr"), DropDownList).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("TRD_EXPIRY_DATE"), HiddenField).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("TRD_MANU_DATE"), HiddenField).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_org_qty2"), TextBox).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("TRD_DRUM_ID_FR"), TextBox).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("TRD_DRUM_LV_FR"), TextBox).ClientID) & "'," &
                                              "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("TRD_SERIAL"), TextBox).ClientID) & "')")

                        tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                        tImage.Attributes.Add("onclick", "LocLookUp(2,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_trd_loc_to"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("trd_loc_to"), HiddenField).ClientID) & "','T')")
                End Select


                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If


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
            Dim seq_string As String = "select MAX(CAST(trd_SEQ AS int)) + 1 from WMS_STOCK_ISSUE_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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


        If TR_WH_FR.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_WH_FR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_WH_FR.Text & "不能空白!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        If TR_WH_TO.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_WH_TO.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_WH_TO.Text & "不能空白!", Session("gLang"))
            End If
            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        If TR_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TR_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TR_DATE.Text & "不能空白!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        ElseIf Not gU.isValidDate(TR_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_TR_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_TR_DATE.Text & "!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        Dim itemCount As Integer = 0


        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("mflag"), HiddenField).Value <> "D" Then
                    itemCount += 1


                    Dim frloc As String = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                    Dim toloc As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    Select Case Session("PAGE_SESSION_MENU_CODE")
                        Case "OP_STF"
                            If frloc <> "" And toloc <> "" Then
                                'If frloc = toloc Then
                                '    If Session("gLang") = "E" Then
                                '        uiFun.displayMsg(Me, "", "Transfer Location Cannot Be Same!", Session("gLang"))
                                '    Else
                                '        uiFun.displayMsg(Me, "", "轉移不能一樣!", Session("gLang"))
                                '    End If

                                '    reloadHiddenValue(GridView1)

                                '    Return False
                                'End If
                            ElseIf frloc = "" And toloc <> "" And flag = "ISSUE" Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "From Location Cannot Be Empty!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "由位置不能空白!", Session("gLang"))
                                End If

                                reloadHiddenValue(GridView1)

                                Return False
                            ElseIf frloc <> "" And toloc = "" And flag = "RECEIVE" Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "To Location Cannot Be Empty!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "到位置不能空白!", Session("gLang"))
                                End If

                                reloadHiddenValue(GridView1)

                                Return False
                            ElseIf frloc = "" AndAlso toloc = "" Then
                                'If Session("gLang") = "E" Then
                                '    uiFun.displayMsg(Me, "", "From/To Location Cannot Be Empty!", Session("gLang"))
                                'Else
                                '    uiFun.displayMsg(Me, "", "由/到位置不能空白!", Session("gLang"))
                                'End If

                                'reloadHiddenValue(GridView1)
                                'Return False
                            End If
                        Case Else
                            If frloc <> "" And toloc <> "" Then
                                Dim ValidateLOCSQL As String
                                ValidateLOCSQL = "select 1 from WMS_WH_BIN  where WH_CODE+LOC_KEY='" & toloc & "' "
                                Dim dtValidateLOC As DataTable = gDB.getDataTable(ValidateLOCSQL)
                                If dtValidateLOC.Rows.Count > 0 Then
                                    'Continue
                                Else
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "To Location " & toloc & " is not valid!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", "To Location " & toloc & " is not valid!", Session("gLang"))
                                    End If
                                    Return False
                                End If
                                'If frloc = toloc Then
                                '    If Session("gLang") = "E" Then
                                '        uiFun.displayMsg(Me, "", "Transfer Location Cannot Be Same!", Session("gLang"))
                                '    Else
                                '        uiFun.displayMsg(Me, "", "轉移不能一樣!", Session("gLang"))
                                '    End If

                                '    reloadHiddenValue(GridView1)

                                '    Return False
                                'End If
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
                            ElseIf frloc = "" AndAlso toloc = "" Then
                                'If Session("gLang") = "E" Then
                                '    uiFun.displayMsg(Me, "", "From/To Location Cannot Be Empty!", Session("gLang"))
                                'Else
                                '    uiFun.displayMsg(Me, "", "由/到位置不能空白!", Session("gLang"))
                                'End If

                                'reloadHiddenValue(GridView1)
                                'Return False
                            End If

                    End Select


                    If CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text <> "" Then
                        If uiFun.gvValidate(Me, dt, "trd_qty", "Qty",
                                         CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text) = False Then Return False


                        If CDbl(CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text) <= 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Item Qty Cannot Be Zero!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "物料數量不能零!", Session("gLang"))
                            End If

                            reloadHiddenValue(GridView1)

                            Return False
                        End If

                        If CDbl(CType(GridView1.Rows(i).FindControl("trd_qty"), TextBox).Text) > CDbl(CType(GridView1.Rows(i).FindControl("trd_org_qty"), TextBox).Text) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Item Qty can not be greater than Original Qty!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "Item Qty can not be greater than Original Qty!", Session("gLang"))
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


                    If CType(GridView1.Rows(i).FindControl("trd_seq"), TextBox).Text <> "" Then
                        Dim dupSQL As String = ""
                        Dim dupTbl As New DataTable
                        Dim ItemCode As String = CType(GridView1.Rows(i).FindControl("itm_code"), HiddenField).Value
                        Dim BatchNo As String = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), DropDownList).Text
                        Dim WH_Loc As String = CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text

                        Dim WH As String = WH_Loc.Substring(0, 4)
                        Dim LOC As String = WH_Loc.Substring(4, 8)

                        dupSQL = "select 1 col from  WMS_WAVEPICK_RSVD " &
                               "where ITEM_CODE = '" & gU.dbEncode(ItemCode) & "' " &
                               "and STORER_CODE= '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                               "and WH_CODE= '" & gU.dbEncode(WH) & "' " &
                               "and LOT_NO= '" & gU.dbEncode(BatchNo) & "' "

                        dupTbl = gDB.getDataTable(dupSQL)
                        If dupTbl.Rows.Count > 0 Then
                            'Dim ask As MsgBoxResult = MsgBox("The stock is reserved for some DO. Do you still want to relocate ?", MsgBoxStyle.YesNo, "Alert")
                            'If ask = MsgBoxResult.No Then
                            '    Return False
                            'End If

                            'Dim confirmValue As String = Request.Form("The stock is reserved for some DO. Do you still want to relocate ?")
                            'If confirmValue <> "Yes" Then
                            '    Return False
                            'End If
                        End If
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
        For i = 0 To gv.Rows.Count - 1
            'CType(gv.Rows(i).FindControl("trd_batch_no_fr"), TextBox).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), TextBox).Text
            'CType(gv.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
            'CType(gv.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
        Next
    End Sub

    Protected Function save(Optional ByVal flag As String = "") As Boolean
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
                If ViewState("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    Dim docType As String = ""
                    If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" Then
                        docType = "SRL"
                    ElseIf Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
                        docType = "RED"
                    Else
                        docType = "STF"
                    End If


                    nextNo = DB.getDocNo(docType, gConn, transaction)
                    'nextNo = TR_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from wms_stock_transfer " &
                                "where tr_code = '" & gU.dbEncode(nextNo) & "' " &
                                "and imp_code= '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                "and storer_code= '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_stock_transfer (" &
                        "tr_code, imp_code, storer_code, " &
                        "tr_status, tr_date, tr_by, " &
                        "tr_batch_no, tr_ref_no, " &
                        "tr_wh_fr, tr_wh_to, TR_TO_LOC, TR_TO_DRUM, tr_rem, " &
                        "tr_edi_rft_no, TR_TQ_NO, " &
                         "sys_cb, sys_cd, sys_lub, sys_lud)" &
                        "values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(TR_STATUS.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(TR_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_BY.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(TR_BATCH_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_REF_NO.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(TR_WH_FR.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TR_WH_TO.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TR_TO_LOC.Value.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_TO_DRUM.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_REM.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(TR_EDI_RFT_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TR_TQ_NO.Text.Trim)) & "," &
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
                                        itemSQL = "insert into wms_stock_transfer_d (" &
                                                "tr_code, imp_code, storer_code, " &
                                                "trd_seq, itm_code, pack_key, " &
                                                "trd_batch_no_fr, trd_qty, trd_org_qty, trd_loc_fr, " &
                                                "trd_loc_to, trd_rem, " &
                                                "trd_pallet_no_fr, trd_pallet_no_to, trd_batch_no_to," &
                                                "TRD_SERIAL, TRD_UOM, TRD_UOM2, TRD_QTY2,TRD_ORG_QTY2,TRD_EXPIRY_DATE, TRD_MANU_DATE, TRD_ITM_NAME, TRD_DRUM_ID_FR, TRD_DRUM_ID_TO, TRD_DRUM_LV_FR, TRD_DRUM_LV_TO, TRD_REQ_QTY," &
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                "values (" &
                                                gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_seq").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("itm_code").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pack_key").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_org_qty").ToString.Trim, "0")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_rem").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_SERIAL").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_UOM").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_UOM2").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_QTY2").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_ORG_QTY2").ToString.Trim, "0")) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_EXPIRY_DATE").ToString.Trim, ""))) & "," &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_MANU_DATE").ToString.Trim, ""))) & "," &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_ITM_NAME").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_ID_FR").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_LV_FR").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_REQ_QTY").ToString.Trim, "0")) & ", " &
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
                        Return False
                        Exit Function
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_stock_transfer set " &
                                    "tr_status = " & gU.convdbNVCData(gU.dbEncode(TR_STATUS.Text)) & ", " &
                                    "tr_date = " & gU.convdbDate(gU.dbEncode(TR_DATE.Text.Trim)) & ", " &
                                    "tr_by = " & gU.convdbNVCData(gU.dbEncode(TR_BY.Text.Trim)) & ", " &
                                    "tr_batch_no = " & gU.convdbNVCData(gU.dbEncode(TR_BATCH_NO.Text.Trim)) & ", " &
                                    "tr_ref_no = " & gU.convdbNVCData(gU.dbEncode(TR_REF_NO.Text.Trim)) & ", " &
                                    "tr_wh_fr = " & gU.convdbNVCData(gU.dbEncode(TR_WH_FR.SelectedValue)) & ", " &
                                    "tr_wh_to = " & gU.convdbNVCData(gU.dbEncode(TR_WH_TO.SelectedValue)) & ", " &
                                    "TR_TO_LOC = " & gU.convdbNVCData(gU.dbEncode(TR_TO_LOC.Value.Trim)) & ", " &
                                    "TR_TO_DRUM = " & gU.convdbNVCData(gU.dbEncode(TR_TO_DRUM.Text.Trim)) & ", " &
                                    "tr_rem = " & gU.convdbNVCData(gU.dbEncode(TR_REM.Text.Trim)) & ", " &
                                    "tr_edi_rft_no = " & gU.convdbNVCData(gU.dbEncode(TR_EDI_RFT_NO.Text.Trim)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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
                                    itemSQL = "insert into wms_stock_transfer_d (" &
                                                "tr_code, imp_code, storer_code, " &
                                                "trd_seq, itm_code, pack_key, " &
                                                "trd_batch_no_fr, trd_qty, trd_org_qty, trd_loc_fr, " &
                                                "trd_loc_to, trd_rem, " &
                                                "trd_pallet_no_fr, trd_pallet_no_to, trd_batch_no_to," &
                                                "TRD_SERIAL, TRD_UOM, TRD_UOM2, TRD_QTY2,TRD_ORG_QTY2,TRD_EXPIRY_DATE, TRD_MANU_DATE, TRD_ITM_NAME, TRD_DRUM_ID_FR, TRD_DRUM_ID_TO, TRD_DRUM_LV_FR, TRD_DRUM_LV_TO," &
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                "values (" &
                                                gU.convdbNVCData(gU.dbEncode(TR_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_seq").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("itm_code").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pack_key").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_org_qty").ToString.Trim, "0")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_rem").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_SERIAL").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_UOM").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_UOM2").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_QTY2").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_ORG_QTY2").ToString.Trim, "0")) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_EXPIRY_DATE").ToString.Trim, ""))) & "," &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_MANU_DATE").ToString.Trim, ""))) & "," &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_ITM_NAME").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_ID_FR").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_LV_FR").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " &
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_stock_transfer_d " &
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and tr_code = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' " &
                                            "and trd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    'itemSQL = "update wms_stock_transfer_d set " & _
                                    '            "trd_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_seq").ToString.Trim, ""))) & ", " & _
                                    '            "itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                    '            "pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                    '            "trd_batch_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                    '            "trd_batch_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                    '            "trd_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " & _
                                    '            "trd_loc_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " & _
                                    '            "trd_loc_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " & _
                                    '            "trd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_rem").ToString.Trim, ""))) & ", " & _
                                    '            "trd_pallet_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                    '            "trd_pallet_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                    '            "sys_lub = '" & Session("usr_id") & "', " & _
                                    '            "sys_lud = Getdate() " & _
                                    '        "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    '        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    '        "and tr_code = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' " & _
                                    '        "and trd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"\

                                    itemSQL = "update wms_stock_transfer_d set " &
                                                "trd_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_qty").ToString.Trim, "0")) & ", " &
                                                "trd_org_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_org_qty").ToString.Trim, "0")) & ", " &
                                                "trd_loc_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_fr").ToString.Trim, ""))) & ", " &
                                                "trd_loc_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_loc_to").ToString.Trim, ""))) & ", " &
                                                "trd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("trd_rem").ToString.Trim, ""))) & ", " &
                                                "TRD_QTY2 = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_QTY2").ToString.Trim, ""))) & ", " &
                                                "TRD_SERIAL = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_SERIAL").ToString.Trim, ""))) & ", " &
                                                "trd_batch_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("trd_batch_no_fr").ToString.Trim, ""))) & ", " &
                                                "TRD_DRUM_ID_TO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, ""))) & ", " &
                                                "TRD_DRUM_LV_TO=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and tr_code = '" & gU.dbEncode(TR_CODE.Text.Trim) & "' " &
                                            "and trd_seq = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("old_seq").ToString.Trim, "")) & "'"

                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                For Each rows As DataRow In dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                dt.AcceptChanges()


                transaction.Commit()
                successFlag = True
                If ViewState("pagemode") = "N" Then
                    'Session.Remove("pagemode")
                    ViewState("pagemode") = Nothing
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
                If flag = "" Then
                    Dim rmtPost As New RemotePost
                    rmtPost.Url = "STFMain.aspx"
                    rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    rmtPost.Add("TR_CODE", ViewState("TR_CODE"))
                    rmtPost.alertMsg = "Record has been Saved successfully!"
                    rmtPost.Post()
                End If
            End If
        End If
        Return successFlag
    End Function

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
            ViewState("TR_CODE") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        TR_STATUS.ForeColor = Drawing.Color.Black

        If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" OrElse Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
            'TR_WH_TO.Enabled = False
        End If


        If ViewState("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            TR_STATUS.Text = "NEW"
            TR_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select Convert(varchar, wms_stock_transfer.TR_DATE," & DDFORMAT & ") as TR_DATE, WMS_WORK_ORDER.WO_CODE, wms_stock_transfer.* from wms_stock_transfer " &
                        "Left Outer Join WMS_WORK_ORDER ON wms_stock_transfer.tr_code = WMS_WORK_ORDER.WO_TR_CODE " &
                        "AND wms_stock_transfer.imp_CODE = WMS_WORK_ORDER.imp_CODE " &
                        "AND wms_stock_transfer.storer_code = WMS_WORK_ORDER.storer_code " &
                        "where wms_stock_transfer.tr_code = '" & gU.dbEncode(pk_code) & "' " &
                        "and wms_stock_transfer.imp_CODE = '" & Session("IMP_CODE") & "' " &
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
                TR_TO_LOC.Value = dt.Rows(0).Item("TR_TO_LOC").ToString
                TR_TO_DRUM.Text = dt.Rows(0).Item("TR_TO_DRUM").ToString
                dsp_TR_TO_LOC.Text = dt.Rows(0).Item("TR_TO_LOC").ToString
                TR_REM.Text = dt.Rows(0).Item("TR_REM").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
                TR_TQ_NO.Text = dt.Rows(0).Item("TR_TQ_NO").ToString
                TR_EDI_RFT_NO.Text = dt.Rows(0).Item("TR_EDI_RFT_NO").ToString

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

                If dt.Rows(0).Item("WO_CODE").ToString <> "" Then
                    WO_TR.Visible = True
                    WO_TR_CODE.Visible = True
                    WO_TR_CODE.Text = dt.Rows(0).Item("WO_CODE").ToString

                    btnGenWO.Visible = False
                Else
                    WO_TR.Visible = False
                    WO_TR_CODE.Visible = False
                End If


                If TR_BY.Text <> "" Then
                    TR_BY.ReadOnly = True
                    TR_BY.BorderWidth = 0
                    TR_BY.BackColor = Drawing.Color.Transparent
                End If

                If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" OrElse Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
                    btnIssue.Visible = False
                    btnRcv.Visible = False

                    btnPost.Visible = True
                Else
                    btnPost.Visible = False

                    Select Case dt.Rows(0).Item("TR_STATUS").ToString
                        Case "NEW"
                            btnIssue.Enabled = True
                            btnRcv.Enabled = False
                        Case "ISSUED"
                            btnIssue.Enabled = False
                            btnRcv.Enabled = True
                        Case "RECEIVED"
                            btnIssue.Enabled = False
                            btnRcv.Enabled = False

                        Case "CANCELLED"
                            btnIssue.Enabled = False
                            btnRcv.Enabled = False

                    End Select

                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT  WMS_STOCK_TRANSFER_D.IMP_CODE, WMS_STOCK_TRANSFER_D.STORER_CODE, WMS_STOCK_TRANSFER_D.TR_CODE, " &
                    "WMS_STOCK_TRANSFER_D.ITM_CODE, WMS_STOCK_TRANSFER_D.PACK_KEY,WMS_STOCK_TRANSFER_D.TRD_SEQ,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_BATCH_NO_FR as TRD_BATCH_NO, WMS_STOCK_TRANSFER_D.TRD_QTY, WMS_STOCK_TRANSFER_D.TRD_ORG_QTY, WMS_STOCK_TRANSFER_D.TRD_PALLET_NO_FR,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_PALLET_NO_TO, WMS_STOCK_TRANSFER_D.TRD_LOC_FR, WMS_STOCK_TRANSFER_D.TRD_LOC_TO,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_REM, WMS_STOCK_TRANSFER_D.SYS_LUB, WMS_STOCK_TRANSFER_D.SYS_LUD, WMS_STOCK_TRANSFER_D.SYS_CD,  " &
                    "WMS_STOCK_TRANSFER_D.SYS_CB, WMS_STOCK_TRANSFER_D.TRD_BATCH_NO_FR, WMS_STOCK_TRANSFER_D.TRD_BATCH_NO_TO,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_SERIAL, WMS_STOCK_TRANSFER_D.TRD_UOM, WMS_STOCK_TRANSFER_D.TRD_UOM2,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_QTY2, WMS_STOCK_TRANSFER_D.TRD_ORG_QTY2, Convert(varchar,WMS_STOCK_TRANSFER_D.TRD_EXPIRY_DATE," & DDFORMAT & ") as TRD_EXPIRY_DATE,Convert(varchar,WMS_STOCK_TRANSFER_D.TRD_MANU_DATE," & DDFORMAT & ") as TRD_MANU_DATE,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_ITM_NAME, WMS_STOCK_TRANSFER_D.TRD_DRUM_ID_FR, WMS_STOCK_TRANSFER_D.TRD_DRUM_ID_TO,  " &
                    "WMS_STOCK_TRANSFER_D.TRD_DRUM_LV_FR, WMS_STOCK_TRANSFER_D.TRD_DRUM_LV_TO, WMS_STOCK_TRANSFER_D.TRD_REQ_QTY, " &
                    "wms_item.itm_name,wms_item.itm_sku_no, 'U' as mFlag, wms_stock_transfer_d.trd_seq as old_seq, '' as LBS_SL " &
                    "from wms_stock_transfer_d, wms_item " &
                    "where wms_stock_transfer_d.imp_code = wms_item.imp_code " &
                    "and wms_stock_transfer_d.storer_code = wms_item.storer_code " &
                    "and wms_stock_transfer_d.itm_code = wms_item.itm_code " &
                    "and wms_stock_transfer_d.PACK_KEY = wms_item.PACK_KEY " &
                    "and wms_stock_transfer_d.tr_code = '" & gU.dbEncode(pk_code) & "' " &
                    "and wms_stock_transfer_d.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    "and wms_stock_transfer_d.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by convert(int,wms_stock_transfer_d.trd_seq)"
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
        Dim cancelSql As String = "update wms_stock_transfer " &
                         "set TR_status = 'CANCELLED', " &
                         "sys_lub = '" & Session("usr_id") & "', " &
                         "sys_lud = Getdate() " &
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                         "and tr_code = '" & gU.dbEncode(TR_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

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
        Dim tmpSL As String = ""

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            If GridView1.Rows.Count > 0 Then
                If validateAll() Then
                    Call save("Y")

                    For Each rows As DataRow In dt.Rows

                        qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  Convert(varchar,max(ILOC_EXPIRY_DATE)," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, max(ILOC_MANU_DATE)," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) &
                                    "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) &
                                    "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) &
                                    "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) &
                                    "' AND ILOC_BATCH_NO = '" & gU.dbEncode(rows.Item("trd_batch_no_fr").ToString.Trim) &
                                    "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("trd_loc_fr").ToString.Trim.Substring(rows.Item("trd_loc_fr").ToString.Trim.Length - 8)) & "' "
                        '"' AND isnull(ILOC_PALLET_NO,'000') = '" & gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) & "' "

                        qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)

                        If qtyTbl.Rows.Count > 0 Then
                            Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")

                            EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                            MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim


                            If CInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, "0")) > CInt(locQty) Then
                                If transaction IsNot Nothing Then
                                    transaction.Rollback()
                                End If

                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Stock: " & gU.dbEncode(rows.Item("itm_sku_no").ToString.Trim) & ", Transfer Qty cannot be Greater than this Location Stock Qty!!\r\nCurrent Location Qty: " & locQty, Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "轉移數量不能大於這個位置儲存數量!!\r\n現時位置儲存數量: " & locQty, Session("gLang"))
                                End If

                                Exit Sub
                            Else


                                st.STORER_CODE = STORER_CODE.SelectedValue
                                st.ITM_CODE = gU.decodeNull(rows.Item("itm_code").ToString.Trim, "")
                                st.PACK_KEY = gU.decodeNull(rows.Item("pack_key").ToString.Trim, "")

                                tmpSL = st.getSLInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)

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
                                st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim.Substring(rows.Item("trd_loc_fr").ToString.Trim.Length - 8), "")
                                st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")

                                If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                    st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                    st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty2").ToString.Trim, ""), "0")
                                    'st.IOS_UOM2 = gU.decodeNull(rows.Item("TRD_UOM2").ToString.Trim, "")
                                    st.IOS_SL = tmpSL
                                    Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                                End If

                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                st.UpdateStockTrans("OUT", gConn, transaction)
                                st.UpdateStockBalTrans("OUT", gConn, transaction)

                                If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                    st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                End If

                                If Session("PAGE_SESSION_MENU_CODE") <> "OP_RD" Then
                                    st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim.Substring(rows.Item("trd_loc_to").ToString.Trim.Length - 8), "")
                                    st.IO_WH = TR_WH_TO.SelectedValue
                                    'st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_to").ToString.Trim, "")

                                    'st.lO_BATCH_NO = gU.decodeNull(rows.Item("trd_batch_no_to").ToString.Trim, "")

                                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                        st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                        st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty2").ToString.Trim, ""), "0")
                                        st.IOS_SL = tmpSL
                                        st.IOS_UOM2 = gU.decodeNull(rows.Item("TRD_UOM2").ToString.Trim, "")
                                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                                        If gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                            st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "")
                                        Else
                                            st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_FR").ToString.Trim, "")
                                        End If

                                        If gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                            st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "")
                                        Else
                                            st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_FR").ToString.Trim, "")
                                        End If
                                        st.IOS_REDRUM_YN = "N"
                                    End If

                                    st.UpdateStockTrans("IN", gConn, transaction)
                                    st.UpdateStockBalTrans("IN", gConn, transaction)

                                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                        st.UpdateStockSerialTrans("IN", gConn, transaction)
                                        st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                    End If
                                End If
                            End If
                            rows.Item("LBS_SL") = tmpSL
                            rows.AcceptChanges()
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

                    If Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
                        For Each rows As DataRow In dt.Rows

                            qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  Convert(varchar,max(ILOC_EXPIRY_DATE)," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, max(ILOC_MANU_DATE)," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) &
                                        "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) &
                                        "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) &
                                        "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) &
                                        "' AND ILOC_LOC = '" & gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim.Substring(rows.Item("trd_loc_fr").ToString.Trim.Length - 8), "") &
                                        "' AND isnull(ILOC_PALLET_NO,'000') = '" & gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) & "' "

                            qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)

                            If qtyTbl.Rows.Count > 0 Then
                                Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")

                                EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                                MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

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
                                st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")
                                st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim.Substring(rows.Item("trd_loc_to").ToString.Trim.Length - 8), "")
                                st.IO_WH = TR_WH_TO.SelectedValue

                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                    If gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "")
                                    Else
                                        st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_FR").ToString.Trim, "")
                                    End If

                                    If gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "")
                                    Else
                                        st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_FR").ToString.Trim, "")
                                    End If
                                    st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                    st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty2").ToString.Trim, ""), "0")
                                    st.IOS_UOM2 = gU.decodeNull(rows.Item("TRD_UOM2").ToString.Trim, "")
                                    Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)

                                    st.IOS_SL = rows.Item("LBS_SL").ToString.Trim
                                End If

                                st.UpdateStockTrans("IN", gConn, transaction)
                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
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

                    updtSql = "update wms_stock_transfer " &
                                "set TR_status = 'POSTED', " &
                                "sys_lub = '" & Session("usr_id") & "', " &
                                "sys_lud = Getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and tr_code = '" & gU.dbEncode(TR_CODE.Text) & "' "

                    gDB.amendData(updtSql, gConn, transaction)

                    transaction.Commit()

                    TR_STATUS.Text = "POSTED"
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

    Protected Sub btnIssue_Click(sender As Object, e As System.EventArgs) Handles btnIssue.Click
        HalfPost("ISSUE")
    End Sub

    Protected Sub btnRcv_Click(sender As Object, e As System.EventArgs) Handles btnRcv.Click
        HalfPost("RECEIVE")
    End Sub

    Protected Sub HalfPost(ByVal flag As String)
        If validateAll(flag) Then
            Dim successFlag As Boolean = False
            Dim updtSql, qtyString As String
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

                    Dim TQNo As String = TR_TQ_NO.Text.Trim


                    For Each rows As DataRow In dt.Rows
                        qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  Convert(varchar,max(ILOC_EXPIRY_DATE)," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, max(ILOC_MANU_DATE)," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) &
                                    "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) &
                                    "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) &
                                    "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) &
                                    "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("trd_loc_fr").ToString.Trim) &
                                    "' AND isnull(ILOC_PALLET_NO,'000') = '" & gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) & "' " &
                                    " AND isnull(ILOC_BATCH_NO,'')='" & gU.dbEncode(rows.Item("TRD_BATCH_NO_FR").ToString.Trim) & "'"

                        qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)


                        If qtyTbl.Rows.Count > 0 Then
                            Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")
                            Dim OS_QTY As Double = 0
                            EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                            MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim


                            qtyString = "select TQD_QTY - isnull(TQD_POST_QTY,0) as OS_QTY from WMS_STOCK_TRANS_REQ_D WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) &
                                    "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) &
                                    "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) &
                                    "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) &
                                    "' AND TQ_CODE = '" & gU.dbEncode(TR_TQ_NO.Text.ToString.Trim) & "'"

                            qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)
                            If qtyTbl.Rows.Count > 0 Then
                                OS_QTY = qtyTbl.Rows(0).Item("OS_QTY")
                            End If

                            'TR_TQ_NO

                            'If gU.dbEncode(rows.Item("trd_pallet_no_fr").ToString.Trim) = "000" Then
                            '    If CInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, "0")) > CInt(locQty) Then
                            '        qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                            '       "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                            '       "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                            '       "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                            '       "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("trd_loc_fr").ToString.Trim) & _
                            '       "' AND isnull(ILOC_PALLET_NO,'000') = '000' " & _
                            '       " AND ILOC_BATCH_NO=" & gU.convdbNVCData(rows.Item("TRD_BATCH_NO_FR").ToString.Trim)
                            '        qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)
                            '        locQty = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item(0).ToString, "0")
                            '    End If
                            'End If

                            If flag = "ISSUE" Then
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
                                ElseIf CInt(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, "0")) > CInt(OS_QTY) Then
                                    If transaction IsNot Nothing Then
                                        transaction.Rollback()
                                    End If

                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Transfer Qty cannot be Greater than the Requested Qty!!\r\nCurrent outstanding Qty: " & OS_QTY, Session("gLang"))
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
                                    st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")

                                    st.IO_WH = TR_WH_FR.SelectedValue
                                    st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_fr").ToString.Trim, "")

                                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                        st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                        If rows.Item("TRD_QTY2").ToString.Trim <> "" And rows.Item("TRD_QTY2").ToString.Trim <> "0" Then
                                            st.IOS_QTY2 = rows.Item("TRD_QTY2").ToString.Trim
                                        Else
                                            st.IOS_QTY2 = 1
                                        End If
                                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                                    End If

                                    st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                    st.UpdateStockTrans("OUT", gConn, transaction)
                                    st.UpdateStockBalTrans("OUT", gConn, transaction)

                                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                                        st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                        st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                    End If

                                    updtSql = "update WMS_STOCK_TRANS_REQ_D set TQD_POST_QTY= isnull(TQD_POST_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("trd_qty").ToString.Trim, "0")) &
                                              ",sys_lub = '" & Session("usr_id") & "', sys_lud = Getdate() " &
                                              " where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "'and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                              "and tq_code = '" & gU.dbEncode(TQNo) & "' AND ITM_CODE='" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & "' and pack_key='" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & "'"
                                    gDB.amendData(updtSql, gConn, transaction)

                                End If
                            ElseIf flag = "RECEIVE" Then
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
                                st.PALLET_NO = gU.decodeNull(rows.Item("trd_pallet_no_fr").ToString.Trim, "")

                                st.IO_WH = TR_WH_TO.SelectedValue
                                st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim, "")

                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                st.UpdateStockTrans("IN", gConn, transaction)

                                REM For update UpdateStockBalTrans
                                st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                If rows.Item("TRD_QTY2").ToString.Trim <> "" And rows.Item("TRD_QTY2").ToString.Trim <> "0" Then
                                    st.IOS_QTY2 = rows.Item("TRD_QTY2").ToString.Trim
                                Else
                                    st.IOS_QTY2 = 1
                                End If
                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then

                                    If gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "")
                                    End If

                                    If gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "")
                                    End If

                                    st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                                    If rows.Item("TRD_QTY2").ToString.Trim <> "" And rows.Item("TRD_QTY2").ToString.Trim <> "0" Then
                                        st.IOS_QTY2 = rows.Item("TRD_QTY2").ToString.Trim
                                    Else
                                        st.IOS_QTY2 = 1
                                    End If
                                    st.IOS_UOM2 = gU.decodeNull(rows.Item("TRD_UOM2").ToString.Trim, "")
                                    st.UpdateStockSerialTrans("IN", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                End If

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

                    Dim statusText As String = ""


                    If flag = "ISSUE" Then
                        statusText = "ISSUED"
                    ElseIf flag = "RECEIVE" Then
                        statusText = "RECEIVED"
                    End If

                    updtSql = "update wms_stock_transfer " &
                                "set TR_status = '" & statusText & "', " &
                                "sys_lub = '" & Session("usr_id") & "', " &
                                "sys_lud = Getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and tr_code = '" & gU.dbEncode(TR_CODE.Text) & "' "

                    gDB.amendData(updtSql, gConn, transaction)


                    updtSql = "update WMS_STOCK_TRANS_REQ " &
                              "set TQ_status = 'PARTIAL', " &
                              "sys_lub = '" & Session("usr_id") & "', " &
                              "sys_lud = Getdate() " &
                              "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                              "and TQ_code = '" & gU.dbEncode(TQNo) & "' "
                    gDB.amendData(updtSql, gConn, transaction)

                    updtSql = "update WMS_STOCK_TRANS_REQ " &
                              "set TQ_status = 'CLOSED',TQ_CLOSED_DATE=getdate(), sys_lub = '" & Session("usr_id") & "', sys_lud = Getdate() " &
                              "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                              "and TQ_code = '" & gU.dbEncode(TQNo) & "' " &
                              "and exists(select 1 from WMS_STOCK_TRANS_REQ_D where TQD_QTY <= ISNULL(TQD_POST_QTY,0) and WMS_STOCK_TRANS_REQ_D.imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                              "and WMS_STOCK_TRANS_REQ_D.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and WMS_STOCK_TRANS_REQ_D.TQ_code = '" & gU.dbEncode(TQNo) & "') " &
                              "and not exists(select 1 from WMS_STOCK_TRANS_REQ_D where TQD_QTY > ISNULL(TQD_POST_QTY,0) and WMS_STOCK_TRANS_REQ_D.imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                              "and WMS_STOCK_TRANS_REQ_D.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and WMS_STOCK_TRANS_REQ_D.TQ_code = '" & gU.dbEncode(TQNo) & "') "

                    gDB.amendData(updtSql, gConn, transaction)


                    transaction.Commit()
                    successFlag = True
                    TR_STATUS.Text = statusText
                    'ar.sec_write = "N"
                    ar.sec_viewMode = "Y"
                    btnPost.Visible = False
                    setPageCtrlAccess()
                    If flag = "ISSUE" Then
                        exceptionEditList.Add("Image_Loc_LookUp_To")
                        exceptionEditList.Add("saveBtn2")
                        exceptionEditList.Add("saveBtn1")
                        exceptionEditList.Add("btnRcv")
                        btnRcv.Enabled = True
                    End If
                    ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

                    uiFun.displayMsg(Me, "", "The Inter-Stock Tranfer has been " & statusText, Session("gLang"))
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

                load_ModalPopupExtender.Hide()

            End Try

        End If
    End Sub

    Protected Sub addItemtoSTF()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(TRD_SEQ AS int)) + 1 from wms_stock_transfer_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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
            Dim addSQL As String = ""

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim seqKeyList As String = ""
            Dim seqListArray As String()

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListArray = Split(seqList.Value, ", ")

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            If seqListArray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListArray.Count - 1
                    seqKeyList = gU.appendToList(seqKeyList, seqListArray(i))
                Next
            End If

            If seqKeyList <> "" Then
                addSQL = "'" & Replace(seqKeyList, ", ", "', '") & "'"
            Else
                addSQL = "NULL"
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


            'SQLString = ""
            'SQLString += "select wms_item.*, wms_alt_vend_item.* "
            'SQLString += "from wms_item left outer join wms_alt_vend_item on "
            'SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            'SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            'SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            'SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            'SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
            'SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            'SQLString += "and wms_item.itm_code || '_000_' || wms_item.pack_key || '_000_' || ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "
            SQLString = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " &
                      " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " &
                      " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_UOM, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " &
                      " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " &
                      " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE 0 END AS REQ_QTY " &
                      " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                      " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " &
                      " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " &
                      " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                      " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                      " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " &
                      " and Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') in (" & addSQL & ") "

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
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = stf_dt.Rows(i).Item("ITM_SKU_NO")
                dt.Rows(rows_count - 1).Item("PACK_KEY") = stf_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("TRD_LOC_FR") = stf_dt.Rows(i).Item("ILOC_LOC")
                If TR_TO_LOC.Value <> "" Then
                    dt.Rows(rows_count - 1).Item("TRD_LOC_TO") = TR_TO_LOC.Value
                Else
                    dt.Rows(rows_count - 1).Item("TRD_LOC_TO") = stf_dt.Rows(i).Item("ILOC_LOC")
                End If
                dt.Rows(rows_count - 1).Item("TRD_ITM_NAME") = stf_dt.Rows(i).Item("ITM_NAME")
                dt.Rows(rows_count - 1).Item("TRD_PALLET_NO_FR") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_PALLET_NO_TO") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_BATCH_NO_FR") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_BATCH_NO_TO") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_SERIAL") = stf_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_UOM") = stf_dt.Rows(i).Item("ITM_UOM")
                dt.Rows(rows_count - 1).Item("TRD_UOM2") = stf_dt.Rows(i).Item("ILBS_UOM2")
                dt.Rows(rows_count - 1).Item("TRD_QTY2") = stf_dt.Rows(i).Item("ILBS_QTY2")
                dt.Rows(rows_count - 1).Item("TRD_ORG_QTY2") = stf_dt.Rows(i).Item("ILBS_QTY2")

                'dt.Rows(rows_count - 1).Item("ISD_VND_CODE") = stf_dt.Rows(i).Item("VND_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("TRD_MANU_DATE") = stf_dt.Rows(i).Item("ILOC_MANU_DATE")
                dt.Rows(rows_count - 1).Item("TRD_EXPIRY_DATE") = stf_dt.Rows(i).Item("ILOC_EXPIRY_DATE")

                dt.Rows(rows_count - 1).Item("TRD_DRUM_ID_FR") = stf_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                If TR_TO_DRUM.Text <> "" Then
                    dt.Rows(rows_count - 1).Item("TRD_DRUM_ID_TO") = TR_TO_DRUM.Text
                Else
                    dt.Rows(rows_count - 1).Item("TRD_DRUM_ID_TO") = stf_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                End If

                dt.Rows(rows_count - 1).Item("TRD_DRUM_LV_FR") = stf_dt.Rows(i).Item("ILBS_DRUM_LEVEL")
                dt.Rows(rows_count - 1).Item("TRD_DRUM_LV_TO") = stf_dt.Rows(i).Item("ILBS_DRUM_LEVEL")

                dt.Rows(rows_count - 1).Item("TRD_ORG_QTY") = stf_dt.Rows(i).Item("ILOC_BAL_QTY")

                dt.Rows(rows_count - 1).Item("TRD_QTY") = stf_dt.Rows(i).Item("REQ_QTY")


                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next

            STORER_CODE.Enabled = False

            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

            For i = 0 To GridView1.Rows.Count - 1

                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
            Next

        End If
    End Sub

    Protected Sub TR_WH_FR_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TR_WH_FR.SelectedIndexChanged
        If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" OrElse Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
            TR_WH_TO.SelectedValue = TR_WH_FR.SelectedValue
            'TR_WH_TO.Enabled = False
        End If
    End Sub

    Protected Sub TR_WH_TO_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TR_WH_TO.SelectedIndexChanged
        If GridView1.Rows.Count > 0 Then
            Dim toArray As New ArrayList
            For i = 0 To GridView1.Rows.Count - 1

                Dim ToLoc As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
                Dim FromLoc As String = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value


                FromLoc = TR_WH_FR.SelectedValue.ToString() + Right(FromLoc, 8)
                'If FromLoc.Length > 12 Then
                '    FromLoc = FromLoc.Substring(TR_WH_FR.SelectedValue.ToString().Length)
                'End If

                CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value = FromLoc
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = FromLoc

                'Dim ToLoc As String = DataBinder.Eval(e.Row.DataItem, "trd_loc_to").ToString.Trim

                ToLoc = TR_WH_TO.SelectedValue.ToString() + Right(ToLoc, 8)

                ''for only QCFG AND FG01 WAREHOUSE'
                'If TR_WH_FR.SelectedValue = "QCFG" AndAlso TR_WH_TO.SelectedValue = "FG01" Then
                '    Dim ToLoc1 As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                '    ToLoc1 = TR_WH_TO.SelectedValue.ToString() + ToLoc

                '    If ToLoc1.Length > 12 Then
                '        ToLoc1 = ToLoc1.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                '        ToLoc = ToLoc1.Remove(ToLoc1.Length - 2, 2) + "00"
                '    End If
                'End If


                'for only QCFG,FGTM AND FGHD WAREHOUSE'
                'for only QCFG,FGTM,FG01,FGLOT AND FGHD WAREHOUSE'
                If TR_WH_TO.SelectedValue = "QCFG" Then
                    Dim ToLocFG As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    ToLocFG = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocFG.Length >= 12 Then
                        ToLocFG = ToLocFG.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocFG.Remove(ToLocFG.Length - 2, 2) + "FG"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGTM" Then
                    Dim ToLocTM As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "TM"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FG01" Then
                    Dim ToLocTM As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "00"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGLOT" Then
                    Dim ToLocTM As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    ToLocTM = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocTM.Length >= 12 Then
                        ToLocTM = ToLocTM.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocTM.Remove(ToLocTM.Length - 2, 2) + "LT"
                    End If
                ElseIf TR_WH_TO.SelectedValue = "FGHD" Then
                    Dim ToLocHD As String = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value

                    ToLocHD = TR_WH_TO.SelectedValue.ToString() + ToLoc

                    If ToLocHD.Length >= 12 Then
                        ToLocHD = ToLocHD.Substring(TR_WH_TO.SelectedValue.ToString().Length)
                        ToLoc = ToLocHD.Remove(ToLocHD.Length - 2, 2) + "HD"
                    End If
                End If

                CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value = ToLoc
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = ToLoc

                'toArray.Add(CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value)
                'CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value = ""
                'CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = ""

                'ViewState("trd_loc_tor") = toArray

                CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value


                'CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                'CType(GridView1.Rows(i).FindControl("dsp_trd_batch_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), HiddenField).Value
                CType(GridView1.Rows(i).FindControl("trd_batch_no_fr"), DropDownList).SelectedIndex = 0


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

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click
        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value, document.getElementById('" & TR_WH_FR.ClientID & "').value);")

        If Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
            Session("IS_Cable") = "Y"
        Else
            Session("IS_Cable") = ""
        End If

        Select Case Session("PAGE_SESSION_MENU_CODE")
            Case "OP_SRL"
                ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "itemLookup", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value,'',document.getElementById('" & TR_WH_FR.ClientID & "').value);", True)
            Case Else
                ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "itemLookup", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value,document.getElementById('" & TR_WH_FR.ClientID & "').value,'');", True)
        End Select

    End Sub

    Protected Sub addItemFromTQ()
        If TQList.Value <> "" Then
            Dim SQLString As String
            Dim stf_dt As DataTable
            Dim addSQL As String = ""
            Dim tempSQL As String = ""
            Dim rows_count As Integer

            Dim itemListarray As String()
            itemListarray = Split(TQList.Value, ", ")

            For i = 0 To itemListarray.Length - 1
                tempSQL &= "'" & itemListarray(i) & "',"
            Next

            addSQL &= " and concat(WMS_STOCK_TRANS_REQ_D.TQ_CODE,'|',WMS_STOCK_TRANS_REQ_D.TQD_SEQ) in (" & Left(tempSQL, Len(tempSQL) - 1) & ") "

            SQLString = " SELECT WMS_STOCK_TRANS_REQ_D.TQ_CODE,WMS_STOCK_TRANS_REQ_D.TQD_SEQ, WMS_STOCK_TRANS_REQ.TQ_WH_FR, WMS_STOCK_TRANS_REQ.TQ_WH_TO, WMS_STOCK_TRANS_REQ_D.ITM_CODE, " &
                        " WMS_STOCK_TRANS_REQ_D.PACK_KEY, WMS_STOCK_TRANS_REQ_D.TQD_QTY, WMS_STOCK_TRANS_REQ_D.TQD_UOM,  " &
                        " WMS_STOCK_TRANS_REQ_D.TQD_UOM2, WMS_STOCK_TRANS_REQ_D.TQD_QTY2, WMS_STOCK_TRANS_REQ_D.TQD_ITM_NAME, WMS_ITEM.ITM_TYPE, WMS_ITEM.ITM_SKU_NO, " &
                        " WMS_STOCK_TRANS_REQ.TQ_EDI_RFT_NO " &
                        " FROM WMS_STOCK_TRANS_REQ INNER JOIN " &
                        " WMS_STOCK_TRANS_REQ_D ON WMS_STOCK_TRANS_REQ.IMP_CODE = WMS_STOCK_TRANS_REQ_D.IMP_CODE AND  " &
                        " WMS_STOCK_TRANS_REQ.STORER_CODE = WMS_STOCK_TRANS_REQ_D.STORER_CODE AND  " &
                        " WMS_STOCK_TRANS_REQ.TQ_CODE = WMS_STOCK_TRANS_REQ_D.TQ_CODE " &
                        " INNER JOIN WMS_ITEM ON WMS_STOCK_TRANS_REQ_D.IMP_CODE = WMS_ITEM.IMP_CODE AND " &
                        " WMS_STOCK_TRANS_REQ_D.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_STOCK_TRANS_REQ_D.ITM_CODE = WMS_ITEM.ITM_CODE AND " &
                        " WMS_STOCK_TRANS_REQ_D.PACK_KEY = WMS_ITEM.PACK_KEY " &
                        " WHERE 1=1" & addSQL


            stf_dt = gDB.getDataTable(SQLString)

            Dim seq_string As String = "select MAX(CAST(trd_SEQ AS int)) + 1 from WMS_STOCK_TRANSFER_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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

            If Not stf_dt Is Nothing AndAlso stf_dt.Rows.Count > 0 Then
                TR_WH_FR.SelectedValue = stf_dt.Rows(0).Item("TQ_WH_FR").ToString.Trim
                TR_WH_TO.SelectedValue = stf_dt.Rows(0).Item("TQ_WH_TO").ToString.Trim

                TR_EDI_RFT_NO.Text = stf_dt.Rows(0).Item("TQ_EDI_RFT_NO").ToString.Trim
                TR_TQ_NO.Text = stf_dt.Rows(0).Item("TQ_CODE").ToString.Trim

                For i = 0 To stf_dt.Rows.Count - 1
                    If ViewState("n_cur_seq") = "" Then
                        ViewState("n_cur_seq") = next_seq_no
                    Else
                        temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                        ViewState("n_cur_seq") = temp_seq_no.ToString
                    End If

                    dt.Rows.Add()

                    rows_count = dt.Rows.Count

                    dt.Rows(rows_count - 1).Item("TRD_SEQ") = ViewState("n_cur_seq").ToString
                    dt.Rows(rows_count - 1).Item("ITM_CODE") = stf_dt.Rows(i).Item("ITM_CODE")
                    dt.Rows(rows_count - 1).Item("ITM_NAME") = stf_dt.Rows(i).Item("TQD_ITM_NAME")
                    dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = stf_dt.Rows(i).Item("ITM_SKU_NO")
                    dt.Rows(rows_count - 1).Item("PACK_KEY") = stf_dt.Rows(i).Item("PACK_KEY")

                    dt.Rows(rows_count - 1).Item("TRD_ITM_NAME") = stf_dt.Rows(i).Item("TQD_ITM_NAME")
                    'dt.Rows(rows_count - 1).Item("TRD_PALLET_NO_FR") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_PALLET_NO_TO") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_BATCH_NO_FR") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_BATCH_NO_TO") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_SERIAL") = stf_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                    If TR_TO_LOC.Value <> "" Then
                        dt.Rows(rows_count - 1).Item("TRD_LOC_TO") = TR_TO_LOC.Value
                    Else
                        dt.Rows(rows_count - 1).Item("TRD_LOC_TO") = ""
                    End If
                    dt.Rows(rows_count - 1).Item("TRD_UOM") = stf_dt.Rows(i).Item("TQD_UOM")
                    dt.Rows(rows_count - 1).Item("TRD_UOM2") = stf_dt.Rows(i).Item("TQD_UOM2")

                    'dt.Rows(rows_count - 1).Item("ISD_VND_CODE") = stf_dt.Rows(i).Item("VND_CODE").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_MANU_DATE") = stf_dt.Rows(i).Item("ILOC_MANU_DATE")
                    'dt.Rows(rows_count - 1).Item("TRD_EXPIRY_DATE") = stf_dt.Rows(i).Item("ILOC_EXPIRY_DATE")

                    'dt.Rows(rows_count - 1).Item("TRD_DRUM_ID_FR") = stf_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_DRUM_ID_TO") = stf_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("TRD_DRUM_LV_FR") = stf_dt.Rows(i).Item("ILBS_DRUM_LEVEL")
                    'dt.Rows(rows_count - 1).Item("TRD_DRUM_LV_TO") = stf_dt.Rows(i).Item("ILBS_DRUM_LEVEL")

                    'dt.Rows(rows_count - 1).Item("TRD_ORG_QTY") = stf_dt.Rows(i).Item("ILOC_BAL_QTY")

                    dt.Rows(rows_count - 1).Item("TRD_REQ_QTY") = stf_dt.Rows(i).Item("TQD_QTY")

                    If stf_dt.Rows(i).Item("ITM_TYPE").ToString.Trim = "CABLE" Then
                        dt.Rows(rows_count - 1).Item("TRD_QTY") = 1
                        dt.Rows(rows_count - 1).Item("TRD_QTY2") = stf_dt.Rows(i).Item("TQD_QTY2")
                    Else
                        dt.Rows(rows_count - 1).Item("TRD_QTY") = 0
                    End If


                    REM **********************
                    dt.Rows(rows_count - 1).Item("mFlag") = "N"

                Next

                STORER_CODE.Enabled = False

                dt.AcceptChanges()
                ViewState("dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()

                selectSTQ.Visible = False
                For i = 0 To GridView1.Rows.Count - 1

                    CType(GridView1.Rows(i).FindControl("dsp_trd_loc_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_fr"), HiddenField).Value
                    CType(GridView1.Rows(i).FindControl("dsp_trd_pallet_no_fr"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_pallet_no_fr"), HiddenField).Value
                    CType(GridView1.Rows(i).FindControl("dsp_trd_loc_to"), Label).Text = CType(GridView1.Rows(i).FindControl("trd_loc_to"), HiddenField).Value
                Next
            End If
        End If

    End Sub

    Protected Sub btnGenWO_Click(sender As Object, e As System.EventArgs) Handles btnGenWO.Click
        GenWO()
    End Sub

    Protected Sub GenWO()
        Dim DefaultBufferStore As String = "EC000000000000"
        'Dim defaultBagItmCode As String = "RFID_BAG"
        Dim SuccessFlag As Boolean = False
        Dim nextWO As String = ""

        If save("Y") Then

            Dim sqlString As String = ""
            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try


                Dim tempDT As DataTable

                sqlString = " SELECT IMP_CODE, STORER_CODE, TR_CODE, ITM_CODE, PACK_KEY, TRD_SEQ, TRD_QTY, TRD_LOC_FR, TRD_LOC_TO, TRD_UOM, TRD_BATCH_NO_FR, TRD_PALLET_NO_FR" &
                            " FROM WMS_STOCK_TRANSFER_D " &
                            " where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                            " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            " and tr_code = '" & gU.dbEncode(ViewState("TR_CODE")) & "' "

                tempDT = gDB.getDataTable(sqlString, gConn, transaction)


                If tempDT.Rows.Count > 0 Then

                    nextWO = DB.getDocNo("WO", gConn, transaction)

                    sqlString = "INSERT INTO WMS_WORK_ORDER (IMP_CODE, STORER_CODE, WO_CODE, WO_STATUS, WO_DATE, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD, WO_TYPE, WO_REM, WO_TR_CODE) " &
                                "VALUES ('" & gU.dbEncode(IMP_CODE.Value.Trim) & "','" & gU.dbEncode(STORER_CODE.SelectedValue) & "','" & nextWO & "','NEW', getdate()," &
                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate(), 'RFID', NULL,'" & gU.dbEncode(ViewState("TR_CODE")) & "'" &
                                ")"
                    gDB.amendData(sqlString, gConn, transaction)

                    For i = 0 To tempDT.Rows.Count - 1
                        sqlString = " INSERT INTO WMS_WORK_ORDER_D " &
                                    " (IMP_CODE, STORER_CODE, WO_CODE, WOD_SEQ, ITM_CODE, PACK_KEY, WOD_BATCH_NO, WOD_TYPE, WOD_QTY, WOD_PALLET_NO, WOD_LOC, WOD_REM, SYS_LUB, SYS_LUD,  SYS_CB, SYS_CD) values (" &
                                    "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "','" & gU.dbEncode(STORER_CODE.SelectedValue) & "','" & nextWO & "', '" & i + 1 & "', " &
                                    gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("ITM_CODE").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("PACK_KEY").ToString.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("TRD_BATCH_NO_FR").ToString.Trim)) & ",'OUT'," & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("TRD_QTY").ToString.Trim, "NULL")) & "," &
                                    gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("TRD_PALLET_NO_FR").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("TRD_LOC_FR").ToString.Trim)) & ",NULL," &
                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                        gDB.amendData(sqlString, gConn, transaction)
                    Next

                    sqlString = " SELECT distinct BAG.ITM_CODE, BAG.PACK_KEY " &
                                " FROM WMS_STOCK_TRANSFER_D INNER JOIN " &
                                " WMS_ITEM ON WMS_STOCK_TRANSFER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_TRANSFER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " &
                                " WMS_STOCK_TRANSFER_D.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_TRANSFER_D.PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " &
                                " WMS_ITEM AS BAG ON WMS_ITEM.IMP_CODE = BAG.IMP_CODE AND WMS_ITEM.STORER_CODE = BAG.STORER_CODE AND  " &
                                " WMS_ITEM.ITM_PARENT = BAG.ITM_SKU_NO " &
                                " where WMS_STOCK_TRANSFER_D.imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                " and WMS_STOCK_TRANSFER_D.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                " and WMS_STOCK_TRANSFER_D.tr_code = '" & gU.dbEncode(ViewState("TR_CODE")) & "' "
                    Dim ParentDT As DataTable
                    ParentDT = gDB.getDataTable(sqlString, gConn, transaction)

                    If ParentDT.Rows.Count > 0 Then
                        For i = 0 To ParentDT.Rows.Count - 1
                            sqlString = " INSERT INTO WMS_WORK_ORDER_D " &
                                    " (IMP_CODE, STORER_CODE, WO_CODE, WOD_SEQ, ITM_CODE, PACK_KEY, WOD_BATCH_NO, WOD_TYPE, WOD_QTY, WOD_PALLET_NO, WOD_LOC, WOD_REM, SYS_LUB, SYS_LUD,  SYS_CB, SYS_CD) values (" &
                                    "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "','" & gU.dbEncode(STORER_CODE.SelectedValue) & "','" & nextWO & "', '" & tempDT.Rows.Count + i + 1 & "', " &
                                    gU.convdbNVCData(gU.dbEncode(ParentDT.Rows(i).Item("ITM_CODE").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ParentDT.Rows(i).Item("PACK_KEY").ToString.Trim)) & ",NULL,'IN',NULL,NULL," & gU.convdbNVCData(gU.dbEncode(DefaultBufferStore)) & ",NULL," &
                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            gDB.amendData(sqlString, gConn, transaction)
                        Next
                    End If

                    transaction.Commit()
                    SuccessFlag = True
                Else

                    uiFun.displayMsg(Me, "", "Error Encounted! Please try again later.", Session("gLang"))

                End If

                'gDB.amendData(cancelSql, gConn, transaction)



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


        End If

        If SuccessFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "..\WO\WOMain.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("WO_CODE", nextWO)
            rmtPost.Add("frSTF", "Y")
            rmtPost.alertMsg = "Work Order has been Created!"
            rmtPost.Post()
        End If

        'uiFun.displayMsg(Me, "", "The Inter-Stock Tranfer has been", Session("gLang"))
    End Sub

    Protected Sub WO_TR_CODE_Click(sender As Object, e As System.EventArgs) Handles WO_TR_CODE.Click
        Dim rmtPost As New RemotePost
        rmtPost.Url = "..\WO\WOMain.aspx"
        rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
        rmtPost.Add("WO_CODE", WO_TR_CODE.Text)
        rmtPost.Add("frSTF", "Y")
        rmtPost.Post()
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged

        If STORER_CODE.SelectedValue = "" Then
        Else
            TR_WH_FR.SelectedValue = "QCFG"
            TR_WH_TO.SelectedValue = "FG01"
        End If
    End Sub

    Protected Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        Dim pk_code As String = ""
        Dim storerCode As String

        storerCode = Server.UrlDecode(Request("STORER_CODE"))
        ViewState("STORER_CODE") = storerCode

        SQLString = "SELECT '0' as trd_seq,WMS_ITEM.ITM_NAME as TRD_ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE,WMS_ITEM.ITM_UOM as TRD_UOM, " &
                  "WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO AS TRD_PALLET_NO_FR, WMS_ITEM_LOC_BAL.ILOC_LOC as TRD_LOC_FR,WMS_ITEM_LOC_BAL.ILOC_LOC as TRD_LOC_TO, " &
                  "CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS TRD_ORG_QTY, " &
                  "CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE 0 END AS TRD_QTY, " &
                  "CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS TRD_REQ_QTY, " &
                  "WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, " &
                  "WMS_ITEM_LOC_BAL.ILOC_RACK , WMS_ITEM_LOC_BAL.ILOC_BIN, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO AS trd_batch_no_fr,  Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_MANU_DATE,103) as TRD_MANU_DATE, " &
                  "Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE,103) as TRD_EXPIRY_DATE, WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " &
                  "ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '')  + '#_#' + ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, " &
                  "Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') AS value, " &
                  "WMS_ITEM.ITM_SKU_NO,  WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO AS TRD_SERIAL, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID as TRD_DRUM_ID_FR,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID as TRD_DRUM_ID_TO, WMS_ITEM_LOC_BAL_S.ILBS_QTY2 AS TRD_QTY2,WMS_ITEM_LOC_BAL_S.ILBS_QTY2 as TRD_ORG_QTY2, " &
                  "WMS_ITEM_LOC_BAL_S.ILBS_UOM2 AS TRD_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL as TRD_DRUM_LV_FR,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL as TRD_DRUM_LV_TO,WMS_ITEM_LOC_BAL_S.ILBS_SL AS trd_rem,'N' as mFlag " &
                  "FROM WMS_ITEM_LOC_BAL INNER JOIN WMS_ITEM ON " &
                  "WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE " &
                  "and WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " &
                  "and WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE " &
                  "and WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY Left outer JOIN  WMS_ITEM_LOC_BAL_S ON " &
                  "WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                  "Where WMS_ITEM_LOC_BAL.PACK_KEY='1' " &
                  "And CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN WMS_ITEM_LOC_BAL_S.ILBS_QTY2 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END > 0 " &
                  "and WMS_ITEM_LOC_BAL.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " &
                  "and ILOC_WH in('" & TR_WH_FR.SelectedValue.ToString.Trim & "')order by ITM_SKU_NO"


        REM **********************
        dt = gDB.getDataTable(SQLString)

        If dt.Rows.Count > 0 Then
            Dim RowCount As Int16 = 1
            For Each row As DataRow In dt.Rows
                row("trd_seq") = RowCount.ToString
                RowCount = RowCount + 1
            Next
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        ViewState("dt") = dt
        GridView1.DataBind()

    End Sub

End Class

