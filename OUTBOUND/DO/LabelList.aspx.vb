Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_DO_LabelList
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable
    Private dt2() As DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me)
            Exit Sub
        End If

        If Not IsPostBack Then
            ViewState("dt2") = Nothing

        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Item Labels"
            'saveBtn1.Text = "OK"
            'saveBtn2.Text = "OK"
            btnPrint.Text = "Print Labels"
        ElseIf Session("gLang") = "C" Then
            lheader.Text = "標籤"
            'saveBtn1.Text = "确定"
            'saveBtn2.Text = "确定"
            btnPrint.Text = "列印標籤"
        End If

        REM **********************

        REM **********************
        REM Additional CSS
        'RT_TYPE.CssClass = "REQUIRED"
        REM **********************

        dt = Session("dt")

        If ViewState("dt2") IsNot Nothing Then
            dt2 = ViewState("dt2")
        Else
            ReDim dt2(dt.Rows.Count - 1)

            For i As Integer = 0 To dt2.Count - 1
                dt2(i) = dt.Clone
            Next

        End If

        If Not IsPostBack Then
            Call BindGV()
        End If

        ar.hideForm(Me, editMode)
    End Sub

    Protected Sub BindGV()

        dt = Session("dt")

        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here

                Call cU.changeGVLabel(oGridViewRow, e, "Seq No.", "編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內號")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名稱")
                Call cU.changeGVLabel(oGridViewRow, e, "Tracking No.", "追查編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "外箱編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Job No.", "Job No.")
                Call cU.changeGVLabel(oGridViewRow, e, "Print Qty.", "列印數量", HorizontalAlign.Right)

                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("dod_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_DISP_SEQ").ToString.Trim
                CType(e.Row.FindControl("dod_itm_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("dod_pack_key"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("dod_itm_desc"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_DESC").ToString.Trim
                CType(e.Row.FindControl("dod_track_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TRACK_NO").ToString.Trim
                CType(e.Row.FindControl("dod_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("dod_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_CARTON_NO").ToString.Trim
                CType(e.Row.FindControl("label_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "label_qty").ToString.Trim

                REM **********************

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Visible = False
                End If

        End Select
    End Sub

    Protected Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim tempDatable As New DataTable
        Dim dt3 As New DataTable

        tempDatable = dt.Copy

        If cU.gfBuildDataTableforGridView(tempDatable, GridView1, True) Then
            For n As Integer = 0 To dt.Rows.Count - 1
                Dim childGridView As GridView = TryCast(GridView1.Rows(n).FindControl("GridView2"), GridView)

                If cU.gfBuildDataTableforGridView(dt2(n), childGridView, True) Then
                    tempDatable.Merge(dt2(n))
                End If


            Next

            Dim Sort_Col As String = "dod_itm_code, dod_seq ASC"
            Dim nextTable As New DataTable

            nextTable = tempDatable.Clone

            Dim foundRows As DataRow() = tempDatable.Select("", Sort_Col)

            For i As Integer = 0 To foundRows.Count - 1
                nextTable.ImportRow(foundRows(i))
            Next

            nextTable.AcceptChanges()

            tempDatable = nextTable

            nextTable.Dispose()

            Dim shemaTbl As DataTable = CreateLableTable()

            Dim colIdx As Integer = 0
            Dim rowIdx As Integer = 0

            Dim itemLoopTimes As Integer = 0

            Dim newrow As DataRow

            If tempDatable.Rows.Count > 0 Then
                newrow = shemaTbl.NewRow()
                shemaTbl.Rows.Add(newrow)
                shemaTbl.AcceptChanges()
            End If

            For x As Integer = 0 To tempDatable.Rows.Count - 1
                itemLoopTimes = gU.decodeEmptyCInt(tempDatable.Rows(x).Item("label_qty").ToString, 0)

                For i As Integer = 1 To itemLoopTimes
                    If colIdx = 6 Then
                        colIdx = 0

                        newrow = shemaTbl.NewRow()
                        shemaTbl.Rows.Add(newrow)
                        shemaTbl.AcceptChanges()

                        rowIdx += 1
                    End If
                    colIdx += 1

                    shemaTbl.Rows(rowIdx).Item("job_" & colIdx.ToString) = gU.decodeNullOrEmpty(tempDatable.Rows(x).Item("dod_ref_no").ToString, "")
                    'shemaTbl.Rows(rowIdx).Item("job_" & colIdx.ToString) = gU.decodeNullOrEmpty(tempDatable.Rows(x).Item("DOD_TRACK_NO").ToString, "")
                    shemaTbl.Rows(rowIdx).Item("itm_" & colIdx.ToString) = gU.decodeNullOrEmpty(tempDatable.Rows(x).Item("dod_itm_code").ToString, "")
                Next
            Next

            If shemaTbl.Rows.Count > 0 Then Session("DO_LBL_SHEMATABLE") = shemaTbl


            tempDatable.Dispose()

            Dim strScript As String = "window.open('./do_lblPrintOut.aspx','do_lbl_printout','menubar=no,scrollbars=yes,resizable=yes,width=1000,height=760,left=5,top=10');"

            If Not Me.ClientScript Is Nothing Then

                If Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType(), "") Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType(), "", strScript, True)
                End If
            End If
        End If

    End Sub

    Private Function CreateLableTable() As DataTable
        Dim lblDataTable As New DataTable()

        Dim lblColumn As DataColumn

        For i As Integer = 1 To 6
            lblColumn = New DataColumn()
            lblColumn.DataType = Type.[GetType]("System.String")
            lblColumn.ColumnName = "job_" & i.ToString
            lblDataTable.Columns.Add(lblColumn)

            lblColumn = New DataColumn()
            lblColumn.DataType = Type.[GetType]("System.String")
            lblColumn.ColumnName = "itm_" & i.ToString
            lblDataTable.Columns.Add(lblColumn)
        Next
      
        Return lblDataTable

    End Function

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim cGridView As New GridView

        For k As Integer = 0 To dt2.Count - 1
            cGridView = DirectCast(GridView1.Rows(k).FindControl("GridView2"), GridView)

            cU.gfBuildDataTableforGridView(dt2(k), cGridView, True)
        Next


        cGridView = DirectCast(GridView1.Rows(e.RowIndex).FindControl("GridView2"), GridView)

        Dim nDiv As New HtmlGenericControl

        nDiv = DirectCast(GridView1.Rows(e.RowIndex).FindControl("Div2"), HtmlGenericControl)

        nDiv.Style("display") = "block"

        Dim MasterRow As DataRow = dt.Rows(e.RowIndex)
        Dim childRow As DataRow = dt2(e.RowIndex).NewRow


        Dim nRow() As DataRow = dt2(e.RowIndex).Select("DOD_ITM_CODE='" & MasterRow.Item("DOD_ITM_CODE").ToString.Trim & "' AND DOD_PACK_KEY='" & MasterRow.Item("DOD_PACK_KEY").ToString.Trim & "'")

        Dim splitRowIdex As Integer = 1

        If nRow.Count > 0 Then
            splitRowIdex = nRow.Count + 1
        End If

        childRow.Item("dod_seq") = "S" & splitRowIdex
        childRow.Item("dod_disp_seq") = splitRowIdex
        childRow.Item("dod_itm_code") = MasterRow.Item("DOD_ITM_CODE").ToString.Trim
        childRow.Item("dod_pack_key") = MasterRow.Item("DOD_PACK_KEY").ToString.Trim

        childRow.Item("dod_itm_desc") = MasterRow.Item("DOD_ITM_DESC").ToString.Trim
        childRow.Item("dod_track_no") = MasterRow.Item("DOD_TRACK_NO").ToString.Trim
        childRow.Item("dod_pallet_no") = MasterRow.Item("DOD_PALLET_NO").ToString.Trim
        childRow.Item("dod_carton_no") = MasterRow.Item("DOD_CARTON_NO").ToString.Trim
        childRow.Item("label_qty") = 1

        dt2(e.RowIndex).Rows.Add(childRow)
        dt2(e.RowIndex).AcceptChanges()

        ViewState("dt2") = dt2

        Dim displayDT As New DataTable

        displayDT = dt2(e.RowIndex).Clone

        Dim nDataRows() As DataRow = dt2(e.RowIndex).Select("DOD_ITM_CODE = '" & MasterRow.Item("DOD_ITM_CODE").ToString.Trim & "' AND DOD_PACK_KEY = '" & MasterRow.Item("DOD_PACK_KEY").ToString.Trim & "'")

        For i As Integer = 0 To nDataRows.Count - 1
            displayDT.ImportRow(nDataRows(i))
        Next

        displayDT.AcceptChanges()

        cGridView.DataSource = displayDT
        cGridView.DataBind()
    End Sub


    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("dod_itm_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_CODE").ToString.Trim

                CType(e.Row.FindControl("dod_itm_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("dod_pack_key"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("dod_itm_desc"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_DESC").ToString.Trim
                CType(e.Row.FindControl("dod_track_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TRACK_NO").ToString.Trim
                CType(e.Row.FindControl("dod_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("dod_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_CARTON_NO").ToString.Trim
                CType(e.Row.FindControl("label_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "label_qty").ToString.Trim

                REM **********************



        End Select
    End Sub
End Class

