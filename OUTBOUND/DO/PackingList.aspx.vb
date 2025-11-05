Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_DO_PackingList
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable

    Private wmsFunc As New WMSFunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim moduleAction As String
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            Exit Sub
        End If

        moduleAction = Request("moduleAction")

        If Session("pagemode") = "N" Then
            'CancelBtn.Visible = False
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Pack List"
            saveBtn1.Text = "OK"
            saveBtn2.Text = "OK"
            newrow.Text = "Add"
            btnReset.OnClientClick = "return confirm(""Confirm to reset all the items from the DO?"");"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "領料清單"
            saveBtn1.Text = "確定"
            saveBtn2.Text = "確定"
            newrow.Text = "新增"
            btnReset.OnClientClick = "return confirm(""確定重置所有的項目?"");"

        End If
        REM **********************

        If Not IsPostBack Then
            dt = Session("_M_OB_DO_TMP_pi_dt")

            ViewState("_M_OB_DO_TMP_pi_dt") = dt
        Else
            dt = ViewState("_M_OB_DO_TMP_pi_dt")
        End If

        If Not IsPostBack Or moduleAction = "RELOADPL" Then

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                reloadPackList()
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
        End If
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Seq", "編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "領料者", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Lot No.", "貨板編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Origin", "Origin", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Vendor Code", "物件號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.", HorizontalAlign.Left)
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Size", "Pack Size", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Qty. per Carton", "Qty. per Carton", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Picked Qty", "數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Total Number", "數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "N.W.", "N.W.", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "G.W.", "G.W.", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "L", "L", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "W", "W", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "H", "H", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Min. Packing", "最少包裝", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Packed By", "領料者")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)

                'Case DataControlRowType.DataRow

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("pad_display_seq"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pad_display_seq").ToString.Trim
                CType(e.Row.FindControl("pad_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_carton_no").ToString.Trim
                CType(e.Row.FindControl("pad_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_pallet_no").ToString.Trim
                CType(e.Row.FindControl("pad_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_ref_no").ToString.Trim
                CType(e.Row.FindControl("pad_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pad_itm_code").ToString.Trim

                CType(e.Row.FindControl("pad_pack_size"), TextBox).Text = gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0)
                CType(e.Row.FindControl("PAD_QTY_PER_CTN"), TextBox).Text = gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "PAD_QTY_PER_CTN").ToString.Trim, 0)
                CType(e.Row.FindControl("pad_pack_key"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "pad_pack_key").ToString.Trim
                CType(e.Row.FindControl("pad_qty"), TextBox).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_qty").ToString.Trim, 0)

                CType(e.Row.FindControl("pad_origin"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_origin").ToString.Trim

                CType(e.Row.FindControl("pad_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pad_batch_no").ToString.Trim
                CType(e.Row.FindControl("pad_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pad_vnd_code").ToString.Trim

                'uiFun.load_dropdown(CType(e.Row.FindControl("PAD_BATCH_NO"), DropDownList), "select dc_date_code from wms_date_code order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("PAD_BATCH_NO"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "PAD_BATCH_NO").ToString.Trim

                CType(e.Row.FindControl("itm_total_qty"), TextBox).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0) * _
                                                                        gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_qty").ToString.Trim, 0)

                CType(e.Row.FindControl("pad_net_weight"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_net_weight").ToString.Trim
                CType(e.Row.FindControl("pad_gross_weight"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_gross_weight").ToString.Trim

                CType(e.Row.FindControl("pad_length"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_length").ToString.Trim
                CType(e.Row.FindControl("pad_width"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_width").ToString.Trim
                CType(e.Row.FindControl("pad_height"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_height").ToString.Trim
                CType(e.Row.FindControl("pad_min_packing"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_min_packing").ToString.Trim
                CType(e.Row.FindControl("pad_pack_by"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pad_pack_by").ToString.Trim

                'CType(e.Row.FindControl("pad_carton_no"), TextBox).Attributes.Add("onblur", "javascript:calTotalQty('" & e.Row.FindControl("pad_carton_no").ClientID & "'," & _
                '                                                                                                        gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0) & "," & _
                '                                                                                                        "'" & e.Row.FindControl("pad_qty").ClientID & "'," & _
                '                                                                                                        "'" & e.Row.FindControl("itm_total_qty").ClientID & "');")

                'CType(e.Row.FindControl("pad_qty"), TextBox).Attributes.Add("onblur", "javascript:calTotalQty('" & e.Row.FindControl("pad_carton_no").ClientID & "'," & _
                '                                                                                                        gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0) & "," & _
                '                                                                                                        "'" & e.Row.FindControl("pad_qty").ClientID & "'," & _
                '                                                                                                        "'" & e.Row.FindControl("itm_total_qty").ClientID & "');")
                CType(e.Row.FindControl("pad_qty"), TextBox).Attributes.Add("onblur", "javascript:calTotalQty('" & e.Row.FindControl("pad_carton_no").ClientID & "'," & _
                                                                                                                       "'" & e.Row.FindControl("pad_pack_size").ClientID & "'," & _
                                                                                                                       "'" & e.Row.FindControl("pad_qty").ClientID & "'," & _
                                                                                                                       "'" & e.Row.FindControl("itm_total_qty").ClientID & "');")
                CType(e.Row.FindControl("pad_pack_size"), TextBox).Attributes.Add("onblur", "javascript:calPickQty('" & e.Row.FindControl("itm_total_qty").ClientID & "'," & _
                                                                                                                 "'" & e.Row.FindControl("pad_pack_size").ClientID & "'," & _
                                                                                                                 "'" & e.Row.FindControl("pad_qty").ClientID & "'," & _
                                                                                                                 gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0) & ", " & _
                                                                                                                 gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_qty").ToString.Trim, 0) & ");")

                CType(e.Row.FindControl("itm_total_qty"), TextBox).Attributes.Add("onblur", "javascript:calPickQty2('" & e.Row.FindControl("itm_total_qty").ClientID & "'," & _
                                                                                                                "'" & e.Row.FindControl("pad_qty").ClientID & "'," & _
                                                                                                                "'" & e.Row.FindControl("pad_pack_size").ClientID & "'," & _
                                                                                                                gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_pack_size").ToString.Trim, 0) * _
                                                                                                                gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "pad_qty").ToString.Trim, 0) & ");")

                CType(e.Row.FindControl("btnSplit"), Button).CommandArgument = DataBinder.Eval(e.Row.DataItem, "PAD_PACK_NO").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "Z" Then
                    For Each iCell As DataControlFieldCell In e.Row.Cells
                        iCell.BackColor = Drawing.Color.Gold
                    Next

                    CType(e.Row.FindControl("btnSplit"), Button).Visible = False
                End If


                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "R" Then
                    e.Row.Visible = False
                ElseIf DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    CType(e.Row.FindControl("pad_display_seq"), Label).Text = ""
                    Call ar.hideGVRow(GridView1, e.Row)
                End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If validateAll() Then
            If cU.gfBuildDataRowsforGridView(dt, dt.Select("mFlag not in ('D','R')"), GridView1, True) Then
                Dim rows_count As Integer = 0
                REM **********************
                REM Modify Here

                Session("_M_OB_DO_TMP_pi_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pi_seq")) + 1)

                dt.Rows.Add()
                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("pad_pack_no") = Session("_M_OB_DO_TMP_pi_seq").ToString
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
        If e.CommandName = "SplitItem" Then
            If validateAll() Then
                If cU.gfBuildDataRowsforGridView(dt, dt.Select("mFlag not in ('D','R')"), GridView1, True) Then
                    For Each r As DataRow In dt.Select("mFlag not in ('D','R')")
                        If r("pad_carton_no").ToString.Contains("-") Then
                            r("sort_col") = r("pad_carton_no").ToString _
                                                .Substring(r("pad_carton_no").ToString.IndexOf("-"), _
                                                           r("pad_carton_no").ToString.Length - r("pad_carton_no").ToString.IndexOf("-")) _
                                                .Replace("-", "") _
                                                .PadLeft(20, "0")
                        Else
                            r("sort_col") = gU.decodeNullOrEmpty(r("pad_carton_no").ToString, "0").PadLeft(20, "0")
                        End If

                        r.AcceptChanges()
                    Next

                    Dim sortRows() As DataRow = dt.Select("", "mFlag, sort_col, pad_display_seq, pad_itm_code,pad_pack_key, pad_batch_no, pad_pack_no desc")
                    Dim seq As Integer = 1

                    For i As Integer = 0 To sortRows.Count - 1
                        If sortRows(i).Item("mFlag") <> "D" AndAlso sortRows(i).Item("mFlag") <> "R" Then
                            sortRows(i).Item("pad_display_seq") = seq
                            sortRows(i).AcceptChanges()

                            seq += 1
                            'ElseIf sortRows(i).Item("mFlag") = "R" Then
                            '    sortRows(i).Item("mFlag") = "D"
                            '    sortRows(i).AcceptChanges()
                        End If
                    Next

                    Dim rows_count As Integer = 0
                    Dim rowNum As Integer = -1
                    Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                    rowNum = CInt(e.CommandArgument)

                    Session("_M_OB_DO_TMP_pi_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pi_seq")) + 1)

                    Dim splitRow As DataRow = dt.NewRow

                    rows_count = dt.Rows.Count

                    Dim srcRow() As DataRow = dt.Select("mFlag not in ('D', 'R') and pad_pack_no = '" & rowNum & "'")

                    Dim tmpID As Integer = gU.decodeEmptyCInt(dt.Compute("MAX(pad_pack_no)", "").ToString(), 0) + 1

                    Dim disSEQ As Double = gU.decodeEmptyCdbl(dt.Compute("max(pad_display_seq)", "mFlag not in ('D', 'R')").ToString, 0) + 1

                    If srcRow.Count > 0 Then
                        splitRow.Item("pad_pack_no") = tmpID
                        splitRow.Item("pad_itm_code") = srcRow(0).Item("pad_itm_code")
                        splitRow.Item("pad_pack_key") = srcRow(0).Item("pad_pack_key")
                        splitRow.Item("pad_pallet_no") = srcRow(0).Item("pad_pallet_no")
                        splitRow.Item("pad_batch_no") = srcRow(0).Item("pad_batch_no")
                        splitRow.Item("pad_pack_by") = srcRow(0).Item("pad_pack_by")
                        splitRow.Item("imp_code") = srcRow(0).Item("imp_code")
                        splitRow.Item("storer_code") = srcRow(0).Item("storer_code")

                        splitRow.Item("DO_CODE") = srcRow(0).Item("DO_CODE")
                        splitRow.Item("pad_origin") = srcRow(0).Item("pad_origin")
                        splitRow.Item("pad_vnd_code") = srcRow(0).Item("pad_vnd_code")
                        splitRow.Item("pad_net_weight") = srcRow(0).Item("pad_net_weight")
                        splitRow.Item("pad_gross_weight") = srcRow(0).Item("pad_gross_weight")
                        splitRow.Item("pad_length") = srcRow(0).Item("pad_length")
                        splitRow.Item("pad_width") = srcRow(0).Item("pad_width")
                        splitRow.Item("pad_height") = srcRow(0).Item("pad_height")

                        splitRow.Item("pad_pack_size") = srcRow(0).Item("pad_pack_size")
                        splitRow.Item("PAD_QTY_PER_CTN") = srcRow(0).Item("PAD_QTY_PER_CTN")

                        splitRow.Item("pad_qty") = 0

                        If splitRow.Item("pad_carton_no").ToString.Contains("-") Then
                            splitRow.Item("sort_col") = splitRow.Item("pad_carton_no").ToString _
                                               .Substring(splitRow.Item("pad_carton_no").ToString.IndexOf("-"), _
                                                           splitRow.Item("pad_carton_no").ToString.Length - splitRow.Item("pad_carton_no").ToString.IndexOf("-")) _
                                               .Replace("-", "") _
                                               .PadLeft(20, "0")
                        Else
                            splitRow.Item("sort_col") = gU.decodeNullOrEmpty(splitRow.Item("pad_carton_no").ToString, "0").PadLeft(20, "0")
                        End If

                        splitRow.Item("pad_display_seq") = disSEQ

                        splitRow.Item("mFlag") = "Z"

                        dt.Rows.Add(splitRow)

                        dt.AcceptChanges()
                    End If

                    ''uiFun.reOrderDetails(dt, "pad_display_seq")

                    'dt.DefaultView.Sort = "mFlag desc, sort_col, pad_itm_code,  pad_pack_key, pad_batch_no, pad_pack_no desc"

                    BindGV()
                End If
            End If
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))

        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Private Function validateAll() As Boolean
        Dim i As Integer
        Dim javaStr As String

        If DO_PACK_LABEL_QTY_1.Text.Trim <> "" Then
            If Not gU.isDecimal(DO_PACK_LABEL_QTY_1.Text.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Invalid number, Label Qty 1!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "無效的數字, Label Qty 1!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If DO_PACK_LABEL_QTY_2.Text.Trim <> "" Then
            If Not gU.isDecimal(DO_PACK_LABEL_QTY_2.Text.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Invalid number, Label Qty 2!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "無效的數字, Label Qty 2!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If DO_PACK_LABEL_QTY_3.Text.Trim <> "" Then
            If Not gU.isDecimal(DO_PACK_LABEL_QTY_3.Text.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Invalid number, Label Qty 3!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "無效的數字, Label Qty 3!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then
                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pad_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Item Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 物件數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pad_min_packing"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Min. Packing!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 最少包裝!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pad_pack_size"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Pack Size!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 包裝大小!", Session("gLang"))
                        End If
                        Return False
                    End If



                    If CType(GridView1.Rows(i).FindControl("pad_pack_size"), TextBox).Text <> 0 Then
                        Dim total As Double = CType(GridView1.Rows(i).FindControl("itm_total_qty"), TextBox).Text
                        Dim pack As Double = CType(GridView1.Rows(i).FindControl("pad_pack_size"), TextBox).Text
                        If total Mod pack <> 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Invalid Pack Size!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "無效的包裝大小!", Session("gLang"))
                            End If
                            Return False
                        End If

                    End If

                End If
            Next
        End If

        Return True
    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        If validateAll() Then
            If cU.gfBuildDataRowsforGridView(dt, dt.Select("mFlag not in ('D','R')"), GridView1, True) Then
                'case when CHARINDEX('-',pd.PAD_CARTON_NO)> 0 then right('00000000000000000000' + replace(SUBSTRING(pd.PAD_CARTON_NO,1,CHARINDEX('-',pd.PAD_CARTON_NO)),'-',''),20) else right('00000000000000000000' + ISNULL(pd.PAD_CARTON_NO,0),20) end
                For Each r As DataRow In dt.Select("mFlag not in ('D','R')")
                    If r("pad_carton_no").ToString.Contains("-") Then
                        r("sort_col") = r("pad_carton_no").ToString _
                                            .Substring(r("pad_carton_no").ToString.IndexOf("-"), _
                                                       r("pad_carton_no").ToString.Length - r("pad_carton_no").ToString.IndexOf("-")) _
                                            .Replace("-", "") _
                                            .PadLeft(20, "0")
                    Else
                        r("sort_col") = gU.decodeNullOrEmpty(r("pad_carton_no").ToString, "0").PadLeft(20, "0")
                    End If

                    If r("mFlag") = "Z" Then r("mFlag") = "N"

                    r.AcceptChanges()
                Next

                Dim sortRows() As DataRow = dt.Select("", "sort_col, pad_display_seq, mFlag, pad_itm_code,  pad_pack_key, pad_batch_no, pad_pack_no desc")
                Dim seq As Integer = 1

                For i As Integer = 0 To sortRows.Count - 1
                    If sortRows(i).Item("mFlag") <> "D" AndAlso sortRows(i).Item("mFlag") <> "R" Then
                        sortRows(i).Item("pad_display_seq") = seq
                        sortRows(i).AcceptChanges()

                        seq += 1
                        'ElseIf sortRows(i).Item("mFlag") = "R" Then
                        '    sortRows(i).Item("mFlag") = "D"
                        '    sortRows(i).AcceptChanges()
                    End If
                Next

                Session("_M_OB_DO_TMP_pi_dt") = dt
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "update1", "window.opener.document.getElementById('DO_PACK_LABEL_QTY_1').value='" & DO_PACK_LABEL_QTY_1.Text & "';" & _
                                                                                    "window.opener.document.getElementById('DO_PACK_LABEL_QTY_2').value='" & DO_PACK_LABEL_QTY_2.Text & "';" & _
                                                                                    "window.opener.document.getElementById('DO_PACK_LABEL_QTY_3').value='" & DO_PACK_LABEL_QTY_3.Text & "';" & _
                                                                                    "window.opener.document.getElementById('DO_PACK_LABEL_1').value='" & DO_PACK_LABEL_1.Text & "';" & _
                                                                                    "window.opener.document.getElementById('DO_PACK_LABEL_2').value='" & DO_PACK_LABEL_2.Text & "';" & _
                                                                                    "window.opener.document.getElementById('DO_PACK_LABEL_3').value='" & DO_PACK_LABEL_3.Text & "';" & _
                                                                                    "window.open('','_self');window.close();", True)
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
                itmKey = dt.Rows(i).Item("pad_itm_code").ToString.Trim & "#_#" & dt.Rows(i).Item("pad_pack_key").ToString.Trim & "#_#" & _
                gU.decodeNullOrEmpty(dt.Rows(i).Item("pad_pallet_no").ToString.Trim, "000") & "#_#" & _
                gU.decodeNullOrEmpty(dt.Rows(i).Item("pad_vnd_code").ToString.Trim, "000")


                If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("pad_qty"), 0), 0)) > 0 Then
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
                If CInt(dt.Rows(delList(i)).Item("PAD_PACK_NO")) = CInt(Session("_M_OB_DO_TMP_pi_seq")) Then
                    Session("_M_OB_DO_TMP_pi_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pi_seq")) - 1)
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
        Dim sort_col As String = "PAD_DISPLAY_SEQ, mFlag "

        IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
        STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))

        DO_PACK_LABEL_1.Text = Server.UrlDecode(Request("DO_PACK_LABEL_1"))
        DO_PACK_LABEL_2.Text = Server.UrlDecode(Request("DO_PACK_LABEL_2"))
        DO_PACK_LABEL_3.Text = Server.UrlDecode(Request("DO_PACK_LABEL_3"))

        If DO_PACK_LABEL_1.Text = "" Then DO_PACK_LABEL_1.Text = "外箱"
        If DO_PACK_LABEL_2.Text = "" Then DO_PACK_LABEL_2.Text = "年份標籤"
        If DO_PACK_LABEL_3.Text = "" Then DO_PACK_LABEL_3.Text = "月份標籤"

        DO_PACK_LABEL_QTY_1.Text = Server.UrlDecode(Request("DO_PACK_LABEL_QTY_1"))
        DO_PACK_LABEL_QTY_2.Text = Server.UrlDecode(Request("DO_PACK_LABEL_QTY_2"))
        DO_PACK_LABEL_QTY_3.Text = Server.UrlDecode(Request("DO_PACK_LABEL_QTY_3"))

        If Session("pagemode") <> "N" Then
            If Request("DO_CODE") IsNot Nothing Then
                DO_CODE.Value = Server.UrlDecode(Request("DO_CODE"))
            Else
                'btnPrint.Visible = False
            End If
        End If

        'delZeroRow()

        sortDt = dt.Copy

        Dim foundRows As DataRow() = sortDt.Select("", sort_col)
        Dim ds As New DataTable

        ds = dt.Clone
        dt.Rows.Clear()
        For i As Integer = 0 To foundRows.Count - 1
            dt.ImportRow(foundRows(i))
        Next

        'dt.AcceptChanges()

        GridView1.DataSource = dt
        GridView1.DataBind()

        If dt.Rows.Count = 0 Then
            saveBtn1.Enabled = False
            saveBtn2.Enabled = False
        End If
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        reloadPackList()

        If dt.Rows.Count > 0 Then
            If Not saveBtn1.Enabled Then
                saveBtn1.Enabled = True
            End If
            If Not saveBtn2.Enabled Then
                saveBtn2.Enabled = True
            End If
        End If
    End Sub

    Private Sub reloadPackList()

        Dim dtldt As DataTable = Session("dt")
        Dim pi_dt As DataTable = Session("_M_OB_DO_TMP_pi_dt")

        Dim nextSEQ As String = ""

        If dtldt IsNot Nothing AndAlso dtldt.Rows.Count > 0 Then

            If pi_dt Is Nothing AndAlso pi_dt.Rows.Count = 0 Then
                pi_dt = New DataTable

                Dim SQLString As String = "select pd.*, pd.pad_pack_key as pad_pack_size, pi.pld_foi_qty, pi.pld_item_qty, 'U' as mflag,  " & _
                  " from wms_do_packing_d pd inner join WMS_DO_PICKLIST_D pi on " & _
                  "pd.imp_code = pi.imp_code " & _
                  "and pd.storer_code = pi.storer_code " & _
                  "and pd.pad_pack_key= pi.pld_pack_key " & _
                  "and pd.pad_itm_code = pi.pld_item_no " & _
                  "and pd.do_code = pi.do_code " & _
                   "and 1 = 0"

                gDB.getDataTable(SQLString, , , pi_dt)
            Else
                For Each piRow As DataRow In pi_dt.Rows
                    piRow.Item("mFlag") = "R"
                Next

                pi_dt.AcceptChanges()
            End If

            If nextSEQ = "" Then
                Dim seq_string As String = "select MAX(CAST(pad_pack_no AS int)) + 1 from wms_do_packing_d " & _
                            "where IMP_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("IMP_CODE"))) & "' " & _
                            "and STORER_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code").ToString)) & "' " & _
                            "and DO_CODE = '" & gU.dbEncode(dtldt.Rows(0).Item("DO_CODE").ToString) & "' "

                nextSEQ = DB.getValueFromSQL(seq_string)

                If nextSEQ Is Nothing OrElse nextSEQ = "" Then
                    nextSEQ = "1"
                End If
            End If

            Dim indx As Integer = 1
            Dim pl_dt As New DataTable
            pl_dt = Session("_M_OB_DO_TMP_pl_dt")

            For Each row As DataRow In pl_dt.Rows
                Dim newRow = pi_dt.NewRow

                newRow.Item("pad_pack_no") = nextSEQ
                newRow.Item("PAD_DISPLAY_SEQ") = indx
                newRow.Item("pad_pallet_no") = row.Item("pld_pallet_no")
                newRow.Item("pad_itm_code") = row.Item("PLD_ITEM_NO")
                newRow.Item("pad_batch_no") = row.Item("pld_batch_no")
                newRow.Item("pad_pack_key") = row.Item("pld_pack_key")
                'newRow.Item("pad_qty") = row.Item("pld_item_qty")
                newRow.Item("pad_qty") = row.Item("pld_foi_qty")


                newRow.Item("pad_pack_by") = Session("usr_id").ToString

                newRow.Item("imp_code") = Server.UrlDecode(Request("IMP_CODE"))
                newRow.Item("storer_code") = Server.UrlDecode(Request("storer_code"))

                newRow.Item("mFlag") = "N"

                Dim dtRows As DataRow() = dtldt.Select("dod_itm_code + '#_#' + dod_pack_key = '" & row.Item("PLD_ITEM_NO").ToString & "#_#" & _
                                                                                                         row.Item("pld_pack_key").ToString & "'")

                If dtRows.Count > 0 Then
                    newRow.Item("DO_CODE") = dtRows(0).Item("DO_CODE")
                    newRow.Item("pad_vnd_code") = dtRows(0).Item("dod_vnd_code")
                    newRow.Item("pad_carton_no") = dtRows(0).Item("dod_carton_no")

                    'Dim itmString As String = "select d.aitm_pcs_per_pack, d.aitm_origin, d.aitm_net_weight, d.aitm_gross_weight, " & _
                    '                        "d.aitm_length, d.aitm_width, d.aitm_hight, m.itm_pcs_per_uom " & _
                    '                    "from wms_item m " & _
                    '                    "left outer join wms_alt_vend_item d on " & _
                    '                     "m.imp_code = d.imp_code " & _
                    '                      "and m.storer_code = d.storer_code " & _
                    '                      "and m.itm_code = d.itm_code " & _
                    '                      "and m.pack_key = d.pack_key " & _
                    '                      "where d.itm_code = '" & gU.dbEncode(dtRows(0).Item("dod_itm_code").ToString.Trim) & "' " & _
                    '                      "and m.storer_code = '" & gU.dbEncode(dtRows(0).Item("storer_code").ToString.Trim) & "' " & _
                    '                      "and m.imp_code = '" & Session("IMP_CODE") & "' " & _
                    '                      "and d.vnd_code = '" & gU.dbEncode(dtRows(0).Item("dod_vnd_code").ToString.Trim) & "' "


                    Dim itmString As String = "select d.aitm_pcs_per_pack, " & _
                                                                "ISNULL(d.aitm_origin, v.aitm_origin) as aitm_origin, " & _
                                                                "ISNULL(d.aitm_net_weight, v.aitm_net_weight) as aitm_net_weight, " & _
                                                                "ISNULL(d.aitm_gross_weight, v.aitm_gross_weight) as aitm_gross_weight, " & _
                                                                "ISNULL(d.aitm_length, v.aitm_length) as aitm_length, " & _
                                                                "ISNULL(d.aitm_width, v.aitm_width) as aitm_width, " & _
                                                                "ISNULL(d.aitm_hight, v.aitm_hight) as aitm_hight, " & _
                                                                "m.itm_pcs_per_uom, " & _
                                                                "ISNULL(d.aitm_qty_per_ctn, v.aitm_qty_per_ctn) as aitm_qty_per_ctn " & _
                                                            "from wms_item m " & _
                                                            "left outer join wms_alt_vend_item d on " & _
                                                             "m.imp_code = d.imp_code " & _
                                                                  "and m.storer_code = d.storer_code " & _
                                                                  "and m.itm_code = d.itm_code " & _
                                                                  "and m.pack_key = d.pack_key " & _
                                                                  "and d.vnd_code = '" & gU.dbEncode(dtRows(0).Item("dod_vnd_code").ToString.Trim) & "' " & _
                                                            "left outer join V_ALT_VEND_ITEM v " & _
                                                             "on m.IMP_CODE = v.IMP_CODE " & _
                                                                 "and m.STORER_CODE = v.STORER_CODE " & _
                                                                 "and m.ITM_CODE = v.ITM_CODE " & _
                                                                 "and m.PACK_KEY = v.PACK_KEY " & _
                                                            "where m.itm_code = '" & gU.dbEncode(dtRows(0).Item("dod_itm_code").ToString.Trim) & "' " & _
                                                            "and m.storer_code = '" & gU.dbEncode(dtRows(0).Item("storer_code").ToString.Trim) & "' " & _
                                                            "and m.imp_code = '" & Session("IMP_CODE") & "' "

                    Dim itmDt As DataTable = gDB.getDataTable(itmString)


                    If itmDt.Rows.Count > 0 Then
                        newRow.Item("pad_net_weight") = itmDt.Rows(0).Item("aitm_net_weight")
                        newRow.Item("pad_gross_weight") = itmDt.Rows(0).Item("aitm_gross_weight")
                        newRow.Item("pad_length") = itmDt.Rows(0).Item("aitm_length")
                        newRow.Item("pad_width") = itmDt.Rows(0).Item("aitm_width")
                        newRow.Item("pad_height") = itmDt.Rows(0).Item("aitm_hight")
                        newRow.Item("pad_origin") = itmDt.Rows(0).Item("aitm_origin")
                        'newRow.Item("pad_pack_size") = itmDt.Rows(0).Item("aitm_pcs_per_pack")
                        newRow.Item("pad_pack_size") = itmDt.Rows(0).Item("itm_pcs_per_uom")
                        newRow.Item("PAD_QTY_PER_CTN") = itmDt.Rows(0).Item("aitm_qty_per_ctn")
                    Else
                        newRow.Item("pad_net_weight") = 0
                        newRow.Item("pad_gross_weight") = 0
                        newRow.Item("pad_length") = 0
                        newRow.Item("pad_width") = 0
                        newRow.Item("pad_height") = 0
                    End If
                Else
                    newRow.Item("pad_net_weight") = 0
                    newRow.Item("pad_gross_weight") = 0
                    newRow.Item("pad_length") = 0
                    newRow.Item("pad_width") = 0
                    newRow.Item("pad_height") = 0
                End If

                nextSEQ = (CInt(nextSEQ) + 1).ToString

                indx += 1

                'Session("_M_OB_DO_TMP_pi_seq") = indx

                pi_dt.Rows.InsertAt(newRow, pi_dt.Rows.Count)
                'pi_dt.Rows.Add(newRow)
            Next

            pi_dt.AcceptChanges()
        End If

        For Each r As DataRow In pi_dt.Select("mFlag not in ('D','R')")
            If r("pad_carton_no").ToString.Contains("-") Then
                r("sort_col") = r("pad_carton_no").ToString _
                                    .Substring(r("pad_carton_no").ToString.IndexOf("-"), _
                                               r("pad_carton_no").ToString.Length - r("pad_carton_no").ToString.IndexOf("-")) _
                                    .Replace("-", "") _
                                    .PadLeft(20, "0")
            Else
                r("sort_col") = gU.decodeNullOrEmpty(r("pad_carton_no").ToString, "0").PadLeft(20, "0")
            End If

            r.AcceptChanges()
        Next

        Dim sortRows() As DataRow = pi_dt.Select("", "sort_col, pad_itm_code,  pad_pack_key, pad_batch_no, pad_pack_no desc")
        Dim seq As Integer = 1

        For i As Integer = 0 To sortRows.Count - 1
            If sortRows(i).Item("mFlag") <> "D" AndAlso sortRows(i).Item("mFlag") <> "R" Then
                sortRows(i).Item("pad_display_seq") = seq
                sortRows(i).AcceptChanges()

                seq += 1
            End If
        Next

        Session("_M_OB_DO_TMP_pi_dt") = pi_dt

        dt = Session("_M_OB_DO_TMP_pi_dt")

        ViewState("_M_OB_DO_TMP_pi_dt") = dt

        BindGV()
    End Sub
End Class
