Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class PackingItem
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable

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
            lheader.Text = "Packing Items"
            saveBtn1.Text = "OK"
            saveBtn2.Text = "OK"
            newrow.Text = "Add"
            'saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "封装物件"
            saveBtn1.Text = "确定"
            saveBtn2.Text = "确定"
            newrow.Text = "新增"
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

        dt = Session("_M_OB_DO_TMP_pi_dt")
        If Not IsPostBack Then
            Call BindGV()
        End If

        If Request("DO_STATUS") = "CANCELLED" Then
            ar.sec_write = "N"
        ElseIf Request("DO_STATUS") = "POSTED" Then
            ar.sec_write = "N"
        End If

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
                Call cU.changeGVLabel(oGridViewRow, e, "Pack No.", "封装编号")
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板编号")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Type", "封装形式")
                Call cU.changeGVLabel(oGridViewRow, e, "Total Packs", "封装总量")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)

                'Case DataControlRowType.DataRow
                '    If gU.decodeNull(DataBinder.Eval(e.Row.DataItem, "mFlag"), "").ToString.Trim = "D" Then
                '        e.Row.Visible = False
                '    End If
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("pad_pack_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PAD_PACK_NO").ToString.Trim
                CType(e.Row.FindControl("pad_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PAD_PALLET_NO").ToString.Trim

                If Session("gLang") = "E" Then
                    uiFun.load_dropdown(CType(e.Row.FindControl("pad_pack_type"), DropDownList), "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'PACK_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(CType(e.Row.FindControl("pad_pack_type"), DropDownList), "select colc_code, colc_chi_value from wms_col_code where colc_tabcol = 'PACK_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_CHI_VALUE", , Session("gSelectLabel"))
                End If
                CType(e.Row.FindControl("pad_pack_type"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "PAD_PACK_TYPE").ToString.Trim

                CType(e.Row.FindControl("pad_totl_packs"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PAD_TOTL_PACKS").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    CType(e.Row.FindControl("pad_pack_no"), TextBox).CssClass = "REQUIRED"
                Else
                    CType(e.Row.FindControl("pad_pack_no"), TextBox).Enabled = False
                End If

                REM **********************

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

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Visible = False
                End If

                'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                '    nButton.Enabled = False
                'End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                Dim rows_count As Integer = 0
                REM **********************
                REM Modify Here

                dt.Rows.Add()
                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                'dt.Rows(rows_count - 1).Item("dod_seq") = Session("n_cur_seq").ToString
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
                dt.AcceptChanges()

                'Session("pi_dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()

                sumTotal()
            End If
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        If dt.Rows(e.RowIndex).Item("mFlag") = "N" Then
            dt.Rows(e.RowIndex).Delete()
        Else
            dt.Rows(e.RowIndex).Item("mFlag") = "D"
        End If
        GridView1.Rows(e.RowIndex).Visible = False
        dt.AcceptChanges()
    End Sub

    Private Function validateAll() As Boolean
        Dim i, j As Integer

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then
                    If CType(GridView1.Rows(i).FindControl("pad_pack_no"), TextBox).Text.Trim = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Pack No. cannot be empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "封装编号不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

                    For j = 0 To GridView1.Rows.Count - 1
                        If i <> j Then
                            If CType(GridView1.Rows(i).FindControl("pad_pack_no"), TextBox).Text.Trim = CType(GridView1.Rows(j).FindControl("pad_pack_no"), TextBox).Text.Trim Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Pack No. cannot be duplicated!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "封装编号不能重复!", Session("gLang"))
                                End If
                                Return False
                            End If
                        End If
                    Next

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pad_totl_packs"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Total Packs!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "无效的数字, 封装总量!", Session("gLang"))
                        End If
                        Return False
                    End If
                End If
            Next
        End If

        Return True
    End Function

    Protected Sub sumTotal()
        Dim i As Integer
        Dim total_pallet, total_carton, total_bin As Long

        total_pallet = 0
        total_carton = 0
        total_bin = 0

        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i).Item("mFlag").ToString <> "D" Then
                Select Case (dt.Rows(i).Item("PAD_PACK_TYPE").ToString.Trim)
                    Case "PALLET"
                        total_pallet = total_pallet + CLng(dt.Rows(i).Item("PAD_TOTL_PACKS").ToString.Trim)
                    Case "CARTON"
                        total_carton = total_carton + CLng(dt.Rows(i).Item("PAD_TOTL_PACKS").ToString.Trim)
                    Case "BIN"
                        total_bin = total_bin + CLng(dt.Rows(i).Item("PAD_TOTL_PACKS").ToString.Trim)
                End Select
            End If
        Next
        Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "update", "<script language=""JavaScript"">updtTotal(" & CStr(total_pallet) & ", " & CStr(total_carton) & ", " & CStr(total_bin) & ");</script>")
    End Sub

    Protected Sub save(Optional ByVal flag As String = "")
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                sumTotal()
                Session("_M_OB_DO_TMP_pi_dt") = dt
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", "<script language=""JavaScript"">opener.reloadPackList();window.close(""_self"");</script>")
            End If
        End If
    End Sub

    Protected Sub BindGV()
        Dim i As Integer

        'Remove empty row
        For i = dt.Rows.Count - 1 To 0 Step -1
            If dt.Rows(i).Item("PAD_PACK_NO").ToString.Trim = "" Then
                dt.Rows(i).Delete()
            End If
        Next
        dt.AcceptChanges()
        'Session("dt") = dt
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

End Class
