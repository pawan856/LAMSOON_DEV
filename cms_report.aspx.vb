Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Xml

Partial Class cms_report
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private RptU As New ReportUtils
    Private ar As New AccessRightUtils
    Private dbFun As New DBfunc

    Private menu_code As String = ""
    Private table_name As String = ""
    Private dt_table As String = ""
    Private rpt_title As String = ""

    Private datasource As New DataTable
    Private dt_datasource As New DataTable
    Private srch_table As New DataTable
    Private col_table As New DataTable

    Private NoAccess As Boolean = False

    Const cms_user As String = "wms_user"
    Const cms_customer As String = "wms_customer"
    Const cms_function As String = "wms_function"

    Const cms_search As String = "wms_search"
    Const cms_search_col As String = "wms_search_col"

    Const cms_col_def As String = "wms_col_def"
    Const cms_col_code As String = "wms_col_code"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            menu_code = Request("menu_code")

            'REM ****************************
            'REM Modify Access Right Here
            ar = New AccessRightUtils(menu_code, Session("usr_id"), Me)

            If menu_code Is Nothing Or menu_code = "" Then
                ar.Force_PageEndCtrlClear(Me)
                Exit Sub
            End If

            NoAccess = ar.hideForm(Me)
            'REM ****************************

            NoAccess = False
            srch_table = ViewState("srch_table")
            table_name = ViewState("table_name")
            col_table = ViewState("col_table")
            rpt_title = ViewState("rpt_title")
            datasource = ViewState("popdt")
            dt_datasource = ViewState("dt_tb")

            Dim saveasDDL As New DropDownList

            saveasDDL.ID = "rpt_save_as"

            If srch_table Is Nothing Then
                Dim srch_sql As String = "select " & cms_search & ".*, " & cms_function & ".fun_eng_name, " & cms_function & ".fun_chi_name " & _
                                         " from " & cms_search & ", " & cms_function & " where " & _
                                             cms_search & ".fun_code = " & cms_function & ".fun_code and " & _
                                             cms_search & ".fun_code = '" & menu_code & "' and " & _
                                             cms_search & ".srch_type = 'REPORT'"

                srch_table = gDB.getDataTable(srch_sql)

                If srch_table.Rows.Count > 0 Then

                    table_name = srch_table.Rows(0).Item("srch_mtable").ToString.Trim.Replace(" ", "")
                    dt_table = srch_table.Rows(0).Item("srch_dtables").ToString.Trim.Replace(" ", "")

                    ViewState("popdt") = Nothing
                    ViewState("table_name") = table_name
                    ViewState("srch_table") = srch_table

                    If srch_table.Rows(0).Item("srch_rpt_col_selection_yn").ToString.Trim = "Y" Then
                        BindGV(table_name, dt_table)
                    Else
                        BindControl(table_name, dt_table)
                    End If
                Else
                    Me.Controls.Clear()
                    'If Not NoAccess Then Response.Write("Error: wms_search define data is missing.")
                    Exit Sub
                End If
            End If

            file_dl.Visible = True

            If srch_table.Rows(0).Item("srch_rpt_col_selection_yn").ToString.Trim = "Y" Then
                col_sel.Visible = True

                sort_tbl.Width = "820px"
                sort_td.Width = "80px"
                lblSort.Width = 80
                sort_blank.Visible = True
            Else
                col_sel.Visible = False

                sort_tbl.Width = "60%"
                sort_td.Width = "180px"
                lblSort.Width = 140
                sort_blank.Visible = False

            End If

            If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                save_as_n_td.Visible = True
                uiFun.load_dropdownBy_ColCode(saveasDDL, "CMS_REPORT.FILE_TYPE", Session("gLang"), , , srch_table.Rows(0).Item("srch_rpt_default_filetype").ToString.Trim)
                save_as_n_td.Controls.Add(saveasDDL)
                status_td.ColSpan = 3
            Else
                save_as_n_td.Visible = False
                status_td.ColSpan = 2
            End If

            If srch_table.Rows(0).Item("srch_rpt_show_preview_yn").ToString.Trim = "Y" Then
                rptform.DefaultButton = "btnPreview"
                preview_tr.Visible = True
            Else
                rptform.DefaultButton = "btnDownload"
                preview_tr.Visible = False
            End If

            REM **********************
            REM Modify Here
            If Session("gLang") = "C" Then
                lblSort.Text = "排序方式"

                lbl_File.Text = "報表內容"
                lbl_file_status.Text = "請點擊下面的按鈕開啟/下載您要求的報表"

                lbl_printout.Text = "打開預覽"

                If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                    lbl_download.Text = "列印/下載為"
                Else
                    lbl_download.Text = "列印/下載"
                End If

                btnDownload.ImageUrl = "images/report_download_chi.gif"
            Else
                lblSort.Text = "Sorted By:"

                lbl_File.Text = "Report Detail"
                lbl_file_status.Text = "Please click the button below to Open / Download the report as your request."

                lbl_printout.Text = "Open Preview"

                If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                    lbl_download.Text = "Print / Download As"
                Else
                    lbl_download.Text = "Print / Download"
                End If

                btnDownload.ImageUrl = "images/report_download.gif"
            End If
            REM **********************

            GenerateSearchTable()
            lblHD.Focus()

        Catch ex As Exception
            Me.Controls.Clear()
            Response.Write(ex.Message)
        End Try
    End Sub

    Private Sub BindControl(ByVal masterTbl As String, ByVal detailTbls As String)
        Dim SQLString As String = ""
        Dim dt_SQLString As String = ""
        Dim fieldname As String = ""

        Dim isMulitMTbl As Boolean = False

        If masterTbl.Contains(",") Then

            isMulitMTbl = True

            Dim msTbl_list As String()
            msTbl_list = masterTbl.Split(",")
            masterTbl = ""
            For x As Integer = 0 To msTbl_list.Length - 1
                If masterTbl <> "" Then masterTbl = masterTbl & ","

                If msTbl_list(x).ToUpper.Trim.ToString <> "" Then
                    masterTbl = masterTbl & "'" & msTbl_list(x).ToUpper.Trim.ToString & "'"
                End If
            Next

            detailTbls = ""
        Else
            If detailTbls.Contains(",") Then
                Dim dTbls_list As String()
                dTbls_list = detailTbls.Split(",")
                detailTbls = ""
                For i As Integer = 0 To dTbls_list.Length - 1
                    If detailTbls <> "" Then detailTbls = detailTbls & ","

                    If dTbls_list(i).ToUpper.Trim.ToString <> "" Then
                        detailTbls = detailTbls & "'" & dTbls_list(i).ToUpper.Trim.ToString & "'"
                    End If
                Next
            Else
                detailTbls = "'" & detailTbls.ToUpper.Trim & "'"
            End If
        End If


        If Session("gLang") = "C" Then
            fieldname = "cold_chi_label"
        Else
            fieldname = "cold_eng_label"
        End If

        If isMulitMTbl Then

            If masterTbl = "" Then masterTbl = "''"

            SQLString = "SELECT " & cms_col_def & ".*, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name in (" & masterTbl & ") and cold_display_seq <> 0 order by cold_display_seq "
        Else
            SQLString = "SELECT " & cms_col_def & ".*, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name = '" & masterTbl & "' and cold_display_seq <> 0 order by cold_display_seq "
        End If

        datasource = gDB.getDataTable(SQLString)


        If detailTbls = "" Then detailTbls = "''"

        dt_SQLString = "SELECT " & cms_col_def & ".*, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name in (" & detailTbls & ") and cold_display_seq <> 0 order by cold_display_seq "
        dt_datasource = gDB.getDataTable(dt_SQLString)

        If Not datasource Is Nothing Then

            ViewState("popdt") = datasource

            If datasource.Rows.Count > 0 Then

                SQLString = "SELECT (table_name || '.' || column_name) as value, cold_eng_label, cold_chi_label from " & cms_col_def & " where fun_code = '" & menu_code & "' and cold_display_seq <> 0 order by cold_display_seq "

                uiFun.load_dropdown(DDL_ORDERBY1, SQLString, "value", fieldname, "", Session("gSelectLabel"))
                uiFun.load_dropdown(DDL_ORDERBY2, SQLString, "value", fieldname, "", Session("gSelectLabel"))
                uiFun.load_dropdown(DDL_ORDERBY3, SQLString, "value", fieldname, "", Session("gSelectLabel"))

                If Not dt_datasource Is Nothing Then

                    ViewState("dt_tb") = dt_datasource

                    If dt_datasource.Rows.Count > 0 Then
                        detailBar.Visible = True
                    End If
                End If
            Else
                Me.Controls.Clear()
                If Not NoAccess Then Response.Write("Error: " & cms_col_def & " define data is missing.")
                Exit Sub
            End If
        Else
            Me.Controls.Clear()
            If Not NoAccess Then Response.Write("Error: " & cms_col_def & " define data is missing.")
            Exit Sub
        End If
    End Sub

    Private Sub BindGV(ByVal masterTbl As String, ByVal detailTbls As String)
        Dim SQLString As String = ""
        Dim dt_SQLString As String = ""
        Dim SCString As String = ""
        Dim fieldname As String = ""

        Dim nLblText As String = ""

        Dim isMulitMTbl As Boolean = False

        If masterTbl.Contains(",") Then

            isMulitMTbl = True

            Dim msTbl_list As String()
            msTbl_list = masterTbl.Split(",")
            masterTbl = ""
            For x As Integer = 0 To msTbl_list.Length - 1
                If masterTbl <> "" Then masterTbl = masterTbl & ","

                If msTbl_list(x).ToUpper.Trim.ToString <> "" Then
                    masterTbl = masterTbl & "'" & msTbl_list(x).ToUpper.Trim.ToString & "'"
                End If
            Next

            nLblText = "Report Data"

            detailTbls = ""
        Else
            If detailTbls.Contains(",") Then
                Dim dTbls_list As String()
                dTbls_list = detailTbls.Split(",")
                detailTbls = ""
                For i As Integer = 0 To dTbls_list.Length - 1
                    If detailTbls <> "" Then detailTbls = detailTbls & ","

                    If dTbls_list(i).ToUpper.Trim.ToString <> "" Then
                        detailTbls = detailTbls & "'" & dTbls_list(i).ToUpper.Trim.ToString & "'"
                    End If
                Next
            Else
                detailTbls = "'" & detailTbls.ToUpper.Trim & "'"
            End If
        End If

        If Session("gLang") = "C" Then
            fieldname = "cold_chi_label"
        Else
            fieldname = "cold_eng_label"
        End If

        'If isMulitMTbl Then

        '    If masterTbl = "" Then masterTbl = "''"

        '    SQLString = "SELECT *, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name in (" & masterTbl & ") and cold_display_seq <> 0 order by cold_display_seq "
        'Else
        '    SQLString = "SELECT *, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name = '" & masterTbl & "' and cold_display_seq <> 0 order by cold_display_seq "
        'End If

        SQLString = "SELECT *, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and cold_section_hd = 'H' and cold_display_seq <> 0 order by cold_display_seq "

        datasource = gDB.getDataTable(SQLString)

        If detailTbls = "" Then detailTbls = "''"

        'dt_SQLString = "SELECT *, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and table_name in (" & detailTbls & ") and cold_display_seq <> 0 order by cold_display_seq "
        dt_SQLString = "SELECT *, 'N' as SUMYN from " & cms_col_def & " where fun_code = '" & menu_code & "' and cold_section_hd = 'D' and cold_display_seq <> 0 order by cold_display_seq "
        dt_datasource = gDB.getDataTable(dt_SQLString)

        If Not datasource Is Nothing Then

            ViewState("popdt") = datasource

            If datasource.Rows.Count > 0 Then
                gvrsList.DataSource = datasource
                gvrsList.DataBind()

                SQLString = "SELECT (table_name + '.' + column_name) as value, cold_eng_label, cold_chi_label from " & cms_col_def & " where fun_code = '" & menu_code & "' and cold_display_seq <> 0 order by cold_display_seq "

                uiFun.load_dropdown(DDL_ORDERBY1, SQLString, "value", fieldname, "", Session("gSelectLabel"))
                uiFun.load_dropdown(DDL_ORDERBY2, SQLString, "value", fieldname, "", Session("gSelectLabel"))
                uiFun.load_dropdown(DDL_ORDERBY3, SQLString, "value", fieldname, "", Session("gSelectLabel"))

                If Session("gLang") = "C" Then
                    If nLblText = "" Then lblHD.Text = "主資料" Else lblHD.Text = nLblText
                    lblDT.Text = "子資料"
                Else
                    If nLblText = "" Then lblHD.Text = "Master Data" Else lblHD.Text = nLblText
                    lblDT.Text = "Detail Data"
                End If

                If Not dt_datasource Is Nothing Then

                    ViewState("dt_tb") = dt_datasource

                    If dt_datasource.Rows.Count > 0 Then
                        detailBar.Visible = True
                        gvrsList2.DataSource = dt_datasource
                        gvrsList2.DataBind()
                    End If
                End If
            Else
                Me.Controls.Clear()
                If Not NoAccess Then Response.Write("Error: " & cms_col_def & " define data is missing.")
                Exit Sub
            End If
        Else
            Me.Controls.Clear()
            If Not NoAccess Then Response.Write("Error: " & cms_col_def & " define data is missing.")
            Exit Sub
        End If
    End Sub

    Private Sub GenerateSearchTable()
        Dim tblSQL As String = "select * from " & cms_search_col & " where fun_code = '" & menu_code & "' and srcl_search_seq <> 0 order by srcl_search_seq"
        Dim tbl As New DataTable
        Dim field As String
        Dim page_title As String

        tbl = gDB.getDataTable(tblSQL)

        ViewState("col_table") = tbl

        If Session("gLang") = "C" Then
            field = "srcl_chi_label"
            page_title = "fun_chi_name"
        Else
            field = "srcl_eng_label"
            page_title = "fun_eng_name"
        End If

        Dim table As New Table

        table.BorderWidth = 0
        table.CellPadding = 1
        table.CellSpacing = 1



        table.Style.Add("width", "60%")
        table.HorizontalAlign = HorizontalAlign.Center

        table.ID = "headerTbl"

        If NoAccess Then table.Enabled = False

        hddiv.Controls.Add(table)

        Dim headerrow As New TableRow()
        Dim headercell As New TableCell()
        headercell.ColumnSpan = 2

        Dim tit_table As New Table
        tit_table.BorderWidth = 0
        tit_table.CellPadding = 0
        tit_table.CellSpacing = 0
        tit_table.Style.Add("align", "center")
        tit_table.Style.Add("width", "100%")

        Dim tit_row As New TableRow()
        Dim tit_cell1 As New TableCell()
        Dim tit_cell2 As New TableCell()

        tit_cell1.CssClass = "TITLE"
        tit_cell1.ColumnSpan = 2
        tit_cell1.Style.Add("width", "100%")

        Dim tit_Label As New Label()
        tit_Label.ID = "lblTitle_N"
        tit_Label.Text = srch_table.Rows(0).Item(page_title).ToString
        tit_Label.Font.Bold = True
        tit_cell1.Controls.Add(tit_Label)

        ViewState("rpt_title") = srch_table.Rows(0).Item(page_title).ToString

        'tit_cell2.Width = 757
        'tit_cell2.CssClass = "TITLE"
        'tit_cell2.Text = "&nbsp;"

        tit_row.Cells.Add(tit_cell1)
        'tit_row.Cells.Add(tit_cell2)
        tit_table.Rows.Add(tit_row)

        headercell.Controls.Add(tit_table)

        headerrow.Cells.Add(headercell)
        table.Rows.Add(headerrow)

        For i As Integer = 0 To tbl.Rows.Count - 1
            Dim row As New TableRow()
            Dim tCell As New TableCell()
            Dim tLabel As New Label()

            tLabel.ID = "lbl" & tbl.Rows(i).Item("cold_tabcol").ToString.Replace(".", "_")
            tLabel.Text = tbl.Rows(i).Item(field).ToString & ":"

            tLabel.Font.Size = 10
            tCell.Controls.Add(tLabel)
            tCell.CssClass = "LabelTD"
            tCell.Style.Add("width", "20%")
            tCell.Wrap = False

            row.Cells.Add(tCell)

            If tbl.Rows(i).Item("srcl_isreadonly").ToString.Trim = "Y" Then
                If tbl.Rows(i).Item("srcl_def_value_or_selected_seq").ToString.Trim = "" Then
                    row.Cells.Add(GenerateCell(tbl.Rows(i)))
                Else
                    Dim objValue As String = ""

                    Dim defValue As String() = tbl.Rows(i).Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue _
                                        Where defKey.Contains("SESSION") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        objValue = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    valueQuery = From defKey In defValue _
                                        Where defKey.Contains("REQUEST") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        objValue = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                        elseCase = False

                        If objValue <> "" Then
                            ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                        End If

                        Exit For
                    Next

                    If elseCase Then objValue = tbl.Rows(i).Item("srcl_def_value_or_selected_seq").ToString

                    Dim cell As New TableCell()

                    Select Case tbl.Rows(i).Item("srcl_edit_style").ToString.Trim
                        Case "T", "L"
                            Dim lbl As New Label

                            lbl.ID = tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                            lbl.Width = gU.decodeEmptyCInt(tbl.Rows(i).Item("srcl_col_width").ToString, "120")

                            If objValue <> "" Then
                                lbl.Text = objValue
                                ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                            Else
                                lbl.Text = ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE")
                            End If

                            cell.Controls.Add(lbl)

                        Case "S"

                            Dim ddl As New DropDownList

                            If tbl.Rows(i).Item("srcl_sql").ToString.Trim = "" Then
                                uiFun.load_dropdownBy_ColCode(ddl, tbl.Rows(i).Item("cold_tabcol").ToString.Trim, Session("gLang"))
                            Else
                                uiFun.load_dropdown(ddl, tbl.Rows(i).Item("srcl_sql").ToString.Trim, , , , Session("gSelectLabel"))
                            End If

                            If objValue <> "" Then
                                Dim lblHD As New HiddenField
                                Dim lbl As New Label

                                lblHD.ID = tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                                lbl.ID = "dispLbl_" & tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                                ddl.SelectedValue = objValue
                                ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue

                                lblHD.Value = ddl.SelectedValue
                                lbl.Text = ddl.SelectedItem.Text

                                rptform.Controls.Add(lblHD)
                                cell.Controls.Add(lbl)

                                ddl.Dispose()
                            Else
                                ddl.SelectedValue = ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE")

                                ddl.ID = tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                                cell.Controls.Add(ddl)
                            End If

                    End Select

                    row.Cells.Add(cell)
                End If
            Else
                row.Cells.Add(GenerateCell(tbl.Rows(i)))
            End If

            If tbl.Rows(i).Item("srcl_edit_style").ToString.Trim <> "H" Then
                table.Rows.Add(row)
            End If
        Next
    End Sub

    Private Function gfBuildDataTableforGridView(ByRef sTable As DataTable, ByVal oGridView As GridView) As DataTable
        Dim intColCount As Integer = sTable.Columns.Count - 1
        Dim rowCollection As GridViewRowCollection = oGridView.Rows
        Dim pageIndex As Integer = oGridView.PageIndex
        Dim iRowIndex As Integer
        Dim cellIndex As Integer = 0
        Dim cloneTable As New DataTable
        Dim isChecked As Boolean = False

        gfBuildDataTableforGridView = Nothing

        cloneTable = sTable.Clone

        Try
            If Not pageIndex = 0 Then
                For i As Integer = 1 To pageIndex
                    iRowIndex = iRowIndex + 10
                Next
            End If

            For Each gridRow As GridViewRow In rowCollection
                isChecked = False

                Dim rowCell As TableCellCollection = gridRow.Cells
                For Each itemCell As TableCell In rowCell
                    For Each ctl As Control In itemCell.Controls
                        If TypeOf ctl Is CheckBox Then

                            Dim ckCtrlCheckBox As CheckBox = DirectCast(ctl, CheckBox)

                            If ckCtrlCheckBox.Enabled = True Then
                                Select Case ckCtrlCheckBox.ID
                                    Case "chkSelect"
                                        If ckCtrlCheckBox.Checked Then
                                            isChecked = True
                                        End If
                                    Case "chkSum"
                                        If ckCtrlCheckBox.Checked Then
                                            cU.gpstrFieldValue = "Y"
                                            cU.set_dataset_value_with_name(sTable, iRowIndex, "SUMYN")
                                            'cU.gpstrFieldValue = sTable.Rows(iRowIndex).Item("COLD_ENG_LABEL").ToString & " (Total)"
                                            'cU.set_dataset_value_with_name(sTable, iRowIndex, "COLD_ENG_LABEL")
                                            'cU.gpstrFieldValue = sTable.Rows(iRowIndex).Item("COLD_CHI_LABEL").ToString & " (合计)"
                                            'cU.set_dataset_value_with_name(sTable, iRowIndex, "COLD_CHI_LABEL")
                                        End If
                                End Select
                            End If
                        ElseIf TypeOf ctl Is TextBox Then
                            Dim tbCtrlTextBox As TextBox = DirectCast(ctl, TextBox)
                            If tbCtrlTextBox.Enabled = True Then
                                cU.gpstrFieldValue = cU.gfRemoveInvalidChar(tbCtrlTextBox.Text)
                                cU.set_dataset_value_with_name(sTable, iRowIndex, tbCtrlTextBox.ID)
                            End If
                        ElseIf TypeOf ctl Is DropDownList Then
                            Dim dplCtrlDropDown As DropDownList = DirectCast(ctl, DropDownList)
                            If dplCtrlDropDown.Enabled = True Then
                                cU.gpstrFieldValue = cU.gfRemoveInvalidChar(dplCtrlDropDown.SelectedValue.ToString)
                                cU.set_dataset_value_with_name(sTable, iRowIndex, dplCtrlDropDown.ID)
                            End If
                        End If
                    Next
                Next

                If isChecked Then cloneTable.ImportRow(sTable.Rows(iRowIndex))
                iRowIndex += 1
            Next

            cloneTable.AcceptChanges()

            Dim Sort_Col As String = "cold_display_seq ASC"
            Dim rptTable As New DataTable

            rptTable = sTable.Clone

            Dim foundRows As DataRow() = cloneTable.Select("", Sort_Col)

            For i As Integer = 0 To foundRows.Count - 1
                rptTable.ImportRow(foundRows(i))
            Next

            rptTable.AcceptChanges()

            cloneTable.Dispose()

            gfBuildDataTableforGridView = rptTable

        Catch ex As Exception
            Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "error", _
            "<script language='javascript'>alert('Error: " & ex.Message & "');</script>")
        End Try
    End Function

    Protected Sub gvrsList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvrsList.DataBound
        Dim oGridView As GridView = DirectCast(sender, GridView)
        Dim chkAll As CheckBox = DirectCast(oGridView.HeaderRow.FindControl("chkSelectAll"), CheckBox)

        chkAll.Attributes("onclick") = "ChgAllChkState(this.checked);"

        Dim ArrayValues As New List(Of String)
        ArrayValues.Add(String.Concat("'", chkAll.ClientID, "'"))

        For Each gvr As GridViewRow In oGridView.Rows
            Dim cb As CheckBox = DirectCast(gvr.FindControl("chkSelect"), CheckBox)

            cb.Attributes("onclick") = "ChgHDState();"
            ArrayValues.Add(String.Concat("'", cb.ClientID, "'"))
        Next

        CBCollection.Text = "<script type=""text/javascript"">" & vbCrLf & _
                        "<!--" & vbCrLf & _
                        String.Concat("var chkArray = new Array(", String.Join(",", ArrayValues.ToArray()), ");") & vbCrLf & _
                        "// -->" & vbCrLf & _
                        "</script>"
    End Sub

    Protected Sub gvrsList2_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvrsList2.DataBound
        Dim oGridView As GridView = DirectCast(sender, GridView)
        Dim chkAll As CheckBox = DirectCast(oGridView.HeaderRow.FindControl("chkSelectAll"), CheckBox)

        chkAll.Attributes("onclick") = "ChgAllChkStateDT(this.checked);"

        Dim ArrayValues As New List(Of String)
        ArrayValues.Add(String.Concat("'", chkAll.ClientID, "'"))

        For Each gvr As GridViewRow In oGridView.Rows
            Dim cb As CheckBox = DirectCast(gvr.FindControl("chkSelect"), CheckBox)

            cb.Attributes("onclick") = "ChgDTState();"
            ArrayValues.Add(String.Concat("'", cb.ClientID, "'"))
        Next

        CBCollectionDT.Text = "<script type=""text/javascript"">" & vbCrLf & _
                        "<!--" & vbCrLf & _
                        String.Concat("var chkArrayDT = new Array(", String.Join(",", ArrayValues.ToArray()), ");") & vbCrLf & _
                        "// -->" & vbCrLf & _
                        "</script>"
    End Sub

    Protected Sub gvrsList_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here

                Call cU.changeCellByIndex(e, 0, "", "", HorizontalAlign.Center)
                Call cU.changeCellByIndex(e, 1, "Show Total", "顯示總和")
                Call cU.changeCellByIndex(e, 2, "Field Name", "欄位名稱")
                Call cU.changeCellByIndex(e, 3, "Sequence", "序列")
                Call cU.changeCellByIndex(e, 4, "Alignment", "對齊")
                Call cU.changeCellByIndex(e, 5, "Column Width", "列寬")
                REM **********************
        End Select
    End Sub

    Protected Sub gvrsList2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList2.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here

                Call cU.changeCellByIndex(e, 0, "", "", HorizontalAlign.Center)
                Call cU.changeCellByIndex(e, 1, "Show Total", "顯示總和")
                Call cU.changeCellByIndex(e, 2, "Field Name", "欄位名稱")
                Call cU.changeCellByIndex(e, 3, "Sequence", "序列")
                Call cU.changeCellByIndex(e, 4, "Alignment", "對齊")
                Call cU.changeCellByIndex(e, 5, "Column Width", "列寬")
                REM **********************
        End Select
    End Sub

    Protected Sub gvrsList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim fieldname As String = "cold_eng_label"

                If Session("gLang") = "C" Then
                    fieldname = "cold_chi_label"
                Else
                    fieldname = "cold_eng_label"
                End If

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl("cold_label"), Label).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl("cold_label"), Label).Text = ""
                End If

                fieldname = "cold_display_seq"

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = ""
                End If

                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onfocus") = "javascript:this.select();"
                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                fieldname = "cold_width"

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = ""
                End If

                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                fieldname = "cold_align"

                Dim fDDL As DropDownList = DirectCast(e.Row.FindControl(fieldname), DropDownList)

                uiFun.load_dropdownBy_ColCode(fDDL, "CMS_COL_DEF.COLD_ALIGN", Session("gLang"))

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    fDDL.SelectedValue = gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim.ToUpper, "L")
                End If

                fieldname = "cold_type"

                Dim fCb As CheckBox = DirectCast(e.Row.FindControl("chkSum"), CheckBox)

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    If DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim.ToUpper = "DECIMAL" Then
                        fCb.Enabled = True
                    End If
                End If

                REM **********************

        End Select
    End Sub

    Protected Sub gvrsList2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList2.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim fieldname As String = "cold_eng_label"

                If Session("gLang") = "C" Then
                    fieldname = "cold_chi_label"
                Else
                    fieldname = "cold_eng_label"
                End If

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl("cold_label"), Label).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl("cold_label"), Label).Text = ""
                End If

                fieldname = "cold_display_seq"

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = ""
                End If

                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onfocus") = "javascript:this.select();"
                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                fieldname = "cold_width"

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim
                Else
                    DirectCast(e.Row.FindControl(fieldname), TextBox).Text = ""
                End If

                DirectCast(e.Row.FindControl(fieldname), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                fieldname = "cold_align"

                Dim fDDL As DropDownList = DirectCast(e.Row.FindControl(fieldname), DropDownList)

                uiFun.load_dropdownBy_ColCode(fDDL, "CMS_COL_DEF.COLD_ALIGN", Session("gLang"))

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    fDDL.SelectedValue = gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim.ToUpper, "L")
                End If

                fieldname = "cold_type"

                Dim fCb As CheckBox = DirectCast(e.Row.FindControl("chkSum"), CheckBox)

                If Not DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim = "" Then
                    If DataBinder.Eval(e.Row.DataItem, fieldname).ToString.Trim.ToUpper = "DECIMAL" Then
                        fCb.Enabled = True
                    End If
                End If

                REM **********************

        End Select
    End Sub

    Private Function generateSQL(ByVal valTable As DataTable, ByRef isSumList As ArrayList) As String
        Dim SQLString As String = ""
        Dim WhereString As String = ""
        Dim GrpByString As String = ""
        Dim sortString As String = ""

        Dim DetailTableList As String = ""
        Dim sortExcludeList As String = ""
        Dim tmpString As String = ""

        Dim isDtTableSelected As Boolean = False
        Dim isGroupBy As Boolean = False

        Dim dTbl_List As String() = Nothing

        For i As Integer = 0 To valTable.Rows.Count - 1

            If srch_table.Rows(0).Item("srch_mtable").ToString.Contains(",") Then
                isDtTableSelected = True
            Else
                If valTable.Rows(i).Item("table_name").ToString <> srch_table.Rows(0).Item("srch_mtable").ToString Then
                    If DetailTableList <> "" Then DetailTableList = DetailTableList & ", "
                    DetailTableList = DetailTableList & valTable.Rows(i).Item("table_name").ToString
                    isDtTableSelected = True
                End If
            End If

            If tmpString <> "" Then tmpString = tmpString & ", "

            Select Case valTable.Rows(i).Item("cold_type").ToString
                Case "DATETIME"
                    tmpString = tmpString & "to_char(" & _
                                                valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & _
                                            ", 'yy/mm/dd') as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

                    If GrpByString <> "" Then GrpByString = GrpByString & ", "
                    GrpByString = GrpByString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString

                Case "NVARCHAR"

                    If valTable.Rows(i).Item("cold_col_tabcol").ToString.Trim <> "" Then
                        Dim valueName As String = "COLC_ENG_VALUE"

                        If Session("gLang") = "C" Then
                            valueName = "COLC_CHI_VALUE"
                        End If

                        tmpString = tmpString & "(SELECT " & valueName & " FROM CMS_COL_CODE WHERE " & _
                                                "COLC_TABCOL = '" & valTable.Rows(i).Item("cold_col_tabcol").ToString.Trim & "' " & _
                                                "AND COLC_CODE = " & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & _
                                                ") as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString
                    Else
                        tmpString = tmpString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & _
                                    " as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

                    End If

                    If GrpByString <> "" Then GrpByString = GrpByString & ", "
                    GrpByString = GrpByString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString

                Case "DECIMAL"
                    If valTable.Rows(i).Item("SUMYN").ToString = "Y" Then
                        'tmpString = tmpString & "SUM(" & _
                        '            valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & ")" & _
                        '            " as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

                        'isGroupBy = True

                        'If sortExcludeList <> "" Then sortExcludeList = sortExcludeList & ", "
                        'sortExcludeList = sortExcludeList & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString

                        isSumList.Add(valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString)
                    End If

                    tmpString = tmpString & "to_char(" & _
                                             valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & _
                                         ") as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

                    If GrpByString <> "" Then GrpByString = GrpByString & ", "
                    GrpByString = GrpByString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString

                Case Else
                    tmpString = tmpString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString & _
                                " as " & valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

                    If GrpByString <> "" Then GrpByString = GrpByString & ", "
                    GrpByString = GrpByString & valTable.Rows(i).Item("table_name").ToString & "." & valTable.Rows(i).Item("column_name").ToString

            End Select
        Next

        If tmpString <> "" Then
            SQLString = SQLString & tmpString
        End If

        If valTable.Rows.Count > 0 Then
            If srch_table.Rows.Count > 0 Then
                If srch_table.Rows(0).Item("srch_rpt_sql").ToString.Trim <> "" Then
                    generateSQL = srch_table.Rows(0).Item("srch_rpt_sql").ToString.Trim
                ElseIf srch_table.Rows(0).Item("srch_sql").ToString.Trim <> "" Then
                    generateSQL = srch_table.Rows(0).Item("srch_sql").ToString.Trim
                Else
                    If isDtTableSelected = False Then
                        generateSQL = "select " & SQLString & " from " & srch_table.Rows(0).Item("srch_mtable").ToString
                    Else
                        'If srch_table.Rows(0).Item("srch_dtables").ToString = "" Then
                        '    generateSQL = "select " & SQLString & " from " & srch_table.Rows(0).Item("srch_mtable").ToString
                        'Else
                        Dim dtTables As String = ""
                        Dim joinTableSC As String = ""

                        Dim frm_table_list As String = ""

                        Dim MatchOne As Boolean = False
                        Dim MatchTwo As Boolean = False

                        If srch_table.Rows(0).Item("srch_dtables").ToString.Contains(",") Then
                            Dim nTemp As String() = srch_table.Rows(0).Item("srch_dtables").ToString.Trim.Replace(" ", "").Split(",")

                            For n As Integer = 0 To nTemp.Length - 1
                                If DetailTableList.Contains(nTemp(n)) Then
                                    If dtTables <> "" Then dtTables = dtTables & ", "
                                    dtTables = dtTables & nTemp(n)
                                End If

                            Next
                        Else
                            dtTables = srch_table.Rows(0).Item("srch_dtables").ToString.Trim.Replace(" ", "")
                        End If

                        If dtTables <> "" Then
                            frm_table_list = srch_table.Rows(0).Item("srch_mtable").ToString & ", " & dtTables
                        Else
                            frm_table_list = srch_table.Rows(0).Item("srch_mtable").ToString
                        End If

                        Dim MatchedTblsName As String = ""
                        Dim LastMatchedTblsName As String = ""

                        If srch_table.Rows(0).Item("srch_join_criteria").ToString.ToUpper.Trim.Contains(",") Then
                            Dim scTemp As String() = srch_table.Rows(0).Item("srch_join_criteria").ToString.ToUpper.Trim.Split(",")
                            dTbl_List = frm_table_list.Replace(" ", "").Split(",")

                            For n As Integer = 0 To scTemp.Length - 1
                                MatchedTblsName = ""
                                For b As Integer = 0 To dTbl_List.Length - 1
                                    If scTemp(n).Contains(dTbl_List(b).Trim) Then
                                        If MatchOne = False Then
                                            MatchOne = True
                                            If srch_table.Rows(0).Item("srch_mtable").ToString.Trim <> dTbl_List(b).Trim Then
                                                MatchedTblsName = dTbl_List(b).Trim
                                            End If
                                        Else
                                            MatchTwo = True
                                            If srch_table.Rows(0).Item("srch_mtable").ToString.Trim <> dTbl_List(b).Trim Then
                                                MatchedTblsName = dTbl_List(b).Trim
                                            End If
                                        End If
                                    End If
                                Next

                                If MatchTwo Then
                                    scTemp(n) = scTemp(n).Trim.Replace("* =", "*=").Replace("= *", "=*")

                                    If LastMatchedTblsName <> "" And LastMatchedTblsName = MatchedTblsName Then
                                        joinTableSC = joinTableSC & " AND " & scTemp(n).Replace("*=", "=")
                                    Else
                                        If scTemp(n).Trim.Contains("*=") Then
                                            joinTableSC = joinTableSC & " LEFT OUTER JOIN " & MatchedTblsName & " ON " & scTemp(n).Replace("*=", "=")
                                        ElseIf scTemp(n).Trim.Contains("=*") Then
                                            joinTableSC = joinTableSC & " RIGHT OUTER JOIN " & MatchedTblsName & " ON " & scTemp(n).Replace("=*", "=")
                                        Else
                                            joinTableSC = joinTableSC & " INNER JOIN " & MatchedTblsName & " ON " & scTemp(n)
                                        End If
                                    End If

                                    LastMatchedTblsName = MatchedTblsName

                                    'If joinTableSC <> "" Then joinTableSC = joinTableSC & " AND "
                                    'joinTableSC = joinTableSC & scTemp(n).Trim
                                End If

                                MatchOne = False
                                MatchTwo = False
                            Next
                        Else
                            MatchedTblsName = ""

                            'If frm_table_list.ToUpper.Trim.Contains(",") Then
                            If srch_table.Rows(0).Item("srch_join_criteria").ToString <> "" Then
                                'joinTableSC = joinTableSC & srch_table.Rows(0).Item("srch_join_criteria").ToString
                                dTbl_List = frm_table_list.Replace(" ", "").Split(",")

                                For b As Integer = 0 To dTbl_List.Length - 1
                                    If srch_table.Rows(0).Item("srch_join_criteria").ToString.Contains(dTbl_List(b).Trim) Then
                                        If MatchOne = False Then
                                            MatchOne = True
                                            If srch_table.Rows(0).Item("srch_mtable").ToString.Trim <> dTbl_List(b).Trim Then
                                                MatchedTblsName = dTbl_List(b).Trim
                                            End If
                                        Else
                                            MatchTwo = True
                                            If srch_table.Rows(0).Item("srch_mtable").ToString.Trim <> dTbl_List(b).Trim Then
                                                MatchedTblsName = dTbl_List(b).Trim
                                            End If
                                        End If
                                    End If
                                Next

                                If MatchTwo Then
                                    If srch_table.Rows(0).Item("srch_join_criteria").ToString.Trim.Contains("*=") Then
                                        joinTableSC = joinTableSC & " LEFT OUTER JOIN " & MatchedTblsName & " ON " & srch_table.Rows(0).Item("srch_join_criteria").ToString.Trim.Replace("*=", "=")
                                    ElseIf srch_table.Rows(0).Item("srch_join_criteria").ToString.Trim.Contains("=*") Then
                                        joinTableSC = joinTableSC & " RIGHT OUTER JOIN " & MatchedTblsName & " ON " & srch_table.Rows(0).Item("srch_join_criteria").ToString.Trim.Replace("=*", "=")
                                    Else
                                        joinTableSC = joinTableSC & " INNER OUTER JOIN " & MatchedTblsName & " ON " & srch_table.Rows(0).Item("srch_join_criteria").ToString.Trim
                                    End If
                                End If
                            Else
                                generateSQL = ""
                                Return generateSQL
                            End If
                            'End If
                        End If

                        'If joinTableSC <> "" Then joinTableSC = "WHERE " & joinTableSC

                        If joinTableSC <> "" Then frm_table_list = srch_table.Rows(0).Item("srch_mtable").ToString.Trim & joinTableSC

                        If frm_table_list <> "" Then
                            'generateSQL = "SELECT " & SQLString & " FROM " & frm_table_list & " " & joinTableSC
                            generateSQL = "SELECT " & SQLString & " FROM " & frm_table_list
                        Else
                            generateSQL = ""
                        End If
                        'End If
                    End If
                End If
            Else
                generateSQL = ""
            End If
        Else
            generateSQL = ""
        End If

        If dTbl_List Is Nothing Then
            ReDim dTbl_List(0)
            dTbl_List(0) = srch_table.Rows(0).Item("srch_mtable").ToString
        End If

        If generateSQL <> "" Then
            WhereString = genSearchSQL()

            If WhereString.ToUpper.Contains(" WHERE ") Then
                generateSQL = generateSQL & WhereString
            Else
                If generateSQL.ToUpper.Contains("WHERE") Then
                    generateSQL = generateSQL & WhereString
                Else
                    generateSQL = generateSQL & " WHERE 1 = 1 " & WhereString
                End If

            End If

            If srch_table.Rows(0).Item("srch_groupby").ToString.Trim.ToUpper <> "" Then
                GrpByString = srch_table.Rows(0).Item("srch_groupby").ToString.Trim.ToUpper
                isGroupBy = True
            Else
                GrpByString = ""
                isGroupBy = False
            End If

            Dim OrderValid As Boolean = False
            Dim v As Integer

            If DDL_ORDERBY1.SelectedValue <> "" Then
                If Not sortExcludeList.ToUpper.Contains(DDL_ORDERBY1.SelectedValue.ToUpper) Then

                    For v = 0 To dTbl_List.Length - 1
                        If DDL_ORDERBY1.SelectedValue.Contains(dTbl_List(v).Trim) Then
                            OrderValid = True
                        End If
                    Next

                    If OrderValid Then
                        If sortString <> "" Then
                            sortString = sortString & ",  " & DDL_ORDERBY1.SelectedValue
                        Else
                            sortString = DDL_ORDERBY1.SelectedValue
                        End If

                        If isGroupBy Then
                            If Not GrpByString.ToUpper.Contains(DDL_ORDERBY1.SelectedValue.ToUpper) Then
                                If GrpByString <> "" Then GrpByString = GrpByString & ", "
                                GrpByString = GrpByString & DDL_ORDERBY1.SelectedValue.ToUpper
                            End If
                        End If
                    Else
                        DDL_ORDERBY1.SelectedIndex = 0
                    End If
                Else
                    DDL_ORDERBY1.SelectedIndex = 0
                End If
            End If

            OrderValid = False

            If DDL_ORDERBY2.SelectedValue <> "" Then
                If Not sortExcludeList.ToUpper.Contains(DDL_ORDERBY2.SelectedValue.ToUpper) Then

                    For v = 0 To dTbl_List.Length - 1
                        If DDL_ORDERBY1.SelectedValue.Contains(dTbl_List(v).Trim) Then
                            OrderValid = True
                        End If
                    Next

                    If OrderValid Then
                        If sortString <> "" Then
                            sortString = sortString & ",  " & DDL_ORDERBY2.SelectedValue
                        Else
                            sortString = DDL_ORDERBY2.SelectedValue
                        End If

                        If isGroupBy Then
                            If Not GrpByString.ToUpper.Contains(DDL_ORDERBY2.SelectedValue.ToUpper) Then
                                If GrpByString <> "" Then GrpByString = GrpByString & ", "
                                GrpByString = GrpByString & DDL_ORDERBY2.SelectedValue.ToUpper
                            End If
                        End If
                    Else
                        DDL_ORDERBY2.SelectedIndex = 0
                    End If
                Else
                    DDL_ORDERBY2.SelectedIndex = 0
                End If
            End If

            OrderValid = False

            If DDL_ORDERBY3.SelectedValue <> "" Then
                If Not sortExcludeList.ToUpper.Contains(DDL_ORDERBY3.SelectedValue.ToUpper) Then

                    For v = 0 To dTbl_List.Length - 1
                        If DDL_ORDERBY1.SelectedValue.Contains(dTbl_List(v).Trim) Then
                            OrderValid = True
                        End If
                    Next

                    If OrderValid Then
                        If sortString <> "" Then
                            sortString = sortString & ",  " & DDL_ORDERBY3.SelectedValue
                        Else
                            sortString = DDL_ORDERBY3.SelectedValue
                        End If

                        If isGroupBy Then
                            If Not GrpByString.ToUpper.Contains(DDL_ORDERBY3.SelectedValue.ToUpper) Then
                                If GrpByString <> "" Then GrpByString = GrpByString & ", "
                                GrpByString = GrpByString & DDL_ORDERBY3.SelectedValue.ToUpper
                            End If
                        End If
                    Else
                        DDL_ORDERBY3.SelectedIndex = 0
                    End If

                Else
                    DDL_ORDERBY3.SelectedIndex = 0
                End If
            End If

            If sortString = "" AndAlso srch_table.Rows(0).Item("srch_default_sorting").ToString.Trim <> "" Then sortString = srch_table.Rows(0).Item("srch_default_sorting").ToString

            If isGroupBy Then
                generateSQL = generateSQL.ToUpper & " GROUP BY " & GrpByString.ToUpper
            End If

            If sortString <> "" Then generateSQL = generateSQL & " ORDER BY " & sortString.ToUpper

        End If

    End Function

    Protected Function genSearchSQL() As String
        Dim symbol As String = ""

        symbol = srch_table.Rows(0).Item("srch_app_sql").ToString

        Dim col_name As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "" '

        If srch_table.Rows(0).Item("srch_sessioncol").ToString <> "" Then
            col_name = table_name & "." & srch_table.Rows(0).Item("srch_sessioncol").ToString
        End If

        Dim FieldName As String = "'"
        Dim ctrlName As String = ""
        Dim oper As String = ""

        For x As Integer = 0 To col_table.Rows.Count - 1
            FieldName = col_table.Rows(x).Item("cold_tabcol").ToString.ToUpper
            ctrlName = FieldName.Replace(".", "_").Trim
            oper = col_table.Rows(x).Item("srcl_operator").ToString.ToUpper

            If col_table.Rows(x).Item("srcl_isreadonly").ToString.Trim = "Y" Then
                If col_table.Rows(x).Item("srcl_def_value_or_selected_seq").ToString.Trim = "" Then
                    GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString, oper, ctrlName, FieldName, symbol)
                Else
                    If col_table.Rows(x).Item("srcl_edit_style").ToString = "S" Then
                        GenerateObjectSQL(SCString, col_table.Rows(x), "H", oper, ctrlName, FieldName, symbol, "S")
                    Else
                        GenerateObjectSQL(SCString, col_table.Rows(x), "L", oper, ctrlName, FieldName, symbol, col_table.Rows(x).Item("srcl_edit_style").ToString.Trim)
                    End If
                End If
            Else
                GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString, oper, ctrlName, FieldName, symbol)
            End If


            'Select Case col_table.Rows(x).Item("srcl_edit_style").ToString
            '    Case "T"
            '        Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

            '        If tmpTextbox.Text <> "" Then
            '            SCString = APPSQL(SCString, FieldName & " " & oper & " N'%" & gU.dbEncode(tmpTextbox.Text) & "%' ", symbol)
            '        End If

            '    Case "D"
            '        If col_table.Rows(x).Item("srcl_date_range_yn").ToString.Trim.ToUpper = "N" Then
            '            Dim tmpDateTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

            '            If tmpDateTextbox.Text <> "" Then
            '                SCString = APPSQL(SCString, FieldName & " >= N'" & gU.dbEncode(tmpDateTextbox.Text) & "' ", symbol)
            '                SCString = APPSQL(SCString, FieldName & " < DATEADD(DAY, 1, N'" & gU.dbEncode(tmpDateTextbox.Text) & "') ", symbol)
            '            End If
            '        Else
            '            Dim tmpFromTextbox As TextBox = DirectCast(hddiv.FindControl("FR_" & ctrlName), TextBox)
            '            Dim tmpToTextbox As TextBox = DirectCast(hddiv.FindControl("TO_" & ctrlName), TextBox)


            '            If tmpFromTextbox.Text <> "" Then
            '                SCString = APPSQL(SCString, FieldName & " >= N'" & gU.dbEncode(tmpFromTextbox.Text) & "' ", symbol)
            '            End If

            '            If tmpToTextbox.Text <> "" Then
            '                SCString = APPSQL(SCString, FieldName & " < DATEADD(DAY, 1, N'" & gU.dbEncode(tmpToTextbox.Text) & "') ", symbol)
            '            End If
            '        End If

            '    Case "C"
            '        Dim tmpCBValue As String = ""

            '        Dim tmpCheckBoxList As New CheckBoxList
            '        tmpCheckBoxList = DirectCast(hddiv.FindControl(ctrlName), CheckBoxList)

            '        If tmpCheckBoxList.SelectedIndex <> -1 Then
            '            SCString = APPSQL(SCString, "UPPER(" & FieldName & ") IN (" & uiFun.getValueFromCB(tmpCheckBoxList) & ") ", symbol)
            '        End If
            '    Case "S"
            '        Dim tmpDDLValue As String = ""

            '        Dim tmpDropDownList As New DropDownList
            '        tmpDropDownList = DirectCast(hddiv.FindControl(ctrlName), DropDownList)

            '        If tmpDropDownList.SelectedValue.ToString <> "" Then
            '            SCString = APPSQL(SCString, FieldName & " = N'" & gU.dbEncode(tmpDropDownList.SelectedItem.Value) & "' ", symbol)
            '        End If

            '    Case Else

            'End Select
        Next

        REM **********************
        If srch_table.Rows(0).Item("srch_sessioncol").ToString <> "" Then
            'SCString = SCString & " AND (" & col_name & " = '' or EXISTS (SELECT 1 FROM " & cms_user & ", " & cms_customer & " where " & col_name & " = " & cms_customer & ".cus_code " & _
            '"and " & cms_user & ".usr_group = " & cms_customer & ".cus_group and (" & cms_user & ".usr_id = '" & Session("usr_id") & "' or " & cms_user & ".usr_manager = '" & Session("usr_id") & "'))) "
            Dim sessionList As String()
            Dim tempSessionString As String = ""
            Dim sessFieldName As String = ""

            If srch_table.Rows(0).Item("srch_sessioncol").ToString.Contains(",") Then
                sessionList = srch_table.Rows(0).Item("srch_sessioncol").ToString.Split(",")

                For j As Integer = 0 To sessionList.Length - 1
                    sessFieldName = ""

                    If table_name <> "" Then
                        sessFieldName = table_name & "." & sessionList(j).ToString.ToUpper.Trim
                    Else
                        sessFieldName = sessionList(j).ToString.ToUpper.Trim
                    End If
                    tempSessionString = tempSessionString & " AND " & sessFieldName & " = N'" & Session(sessionList(j).ToUpper.ToString.Trim) & "'"
                Next
            Else
                If srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim.Contains(".") Then
                    tempSessionString = " AND " & srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim & " = N'" & Session(srch_table.Rows(0).Item("srch_sessioncol").ToUpper.ToString.Trim) & "'"
                Else
                    If table_name <> "" Then
                        sessFieldName = table_name & "." & srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim
                    Else
                        sessFieldName = srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim
                    End If

                    tempSessionString = " AND " & sessFieldName & " = N'" & Session(srch_table.Rows(0).Item("srch_sessioncol").ToUpper.ToString.Trim) & "'"
                End If

            End If

            SCString = SCString & tempSessionString
        End If
        REM **********************

        Return SCString

    End Function

    Private Function APPSQL(ByVal SQL As String, ByVal Clause As String, ByVal SYMBOL As String) As String

        APPSQL = ""

        Select Case SYMBOL
            Case "WHERE"
                If SQL <> "" Then
                    APPSQL = SQL & " AND " & Clause
                Else
                    APPSQL = " WHERE " & Clause
                End If

            Case "AND"
                APPSQL = SQL & " AND " & Clause
            Case Else
                APPSQL = SQL & " AND " & Clause
        End Select
    End Function

    Private Sub generateSumColumn(ByRef Source As DataTable, ByVal sList As ArrayList)

        Dim tit_Row As DataRow = Source.NewRow
        Dim exRow As DataRow = Source.NewRow

        Dim sumDict As New Dictionary(Of String, Double)

        Dim tempvalue As Double

        For i As Integer = 0 To Source.Rows.Count - 1
            For x As Integer = 0 To sList.Count - 1

                tempvalue = 0

                If Not sumDict.ContainsKey(sList(x)) Then
                    sumDict.Add(sList(x), gU.decodeEmptyCdbl(Source.Rows(i).Item(sList(x)).ToString(), "0"))
                Else
                    tempvalue = CDbl(sumDict.Item(sList(x).ToString()))
                    sumDict.Item(sList(x)) = tempvalue + gU.decodeEmptyCdbl(Source.Rows(i).Item(sList(x)).ToString(), "0")
                End If

            Next

            'Source.Columns("").Ordinal
        Next

        For n As Integer = 0 To sList.Count - 1
            If sumDict.ContainsKey(sList(n)) Then
                If Session("gLang") = "C" Then
                    tit_Row.Item(sList(n)) = "合計"
                Else
                    tit_Row.Item(sList(n)) = "Sub Total"
                End If

                exRow.Item(sList(n)) = FormatNumber(sumDict.Item(sList(n)), 2, TriState.True)
            End If
        Next

        If sList.Count > 0 Then
            Source.Rows.Add(tit_Row)
            Source.Rows.Add(exRow)
            Source.AcceptChanges()
        End If
    End Sub

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnDownload.Click
        If srch_table.Rows(0).Item("srch_rpt_col_selection_yn").ToString.Trim = "Y" Then
            Call showReport(True)
        Else
            Call showReport(False)
        End If
    End Sub

    Protected Sub btnPreview_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPreview.Click
        If srch_table.Rows(0).Item("srch_rpt_col_selection_yn").ToString.Trim = "Y" Then
            Call showReport(True, True)
        Else
            Call showReport(False, True)
        End If
    End Sub

    Private Sub showReport(ByVal isColSelection As Boolean, Optional ByVal isPreview As Boolean = False)
        Dim report_table As New DataTable
        Dim dt_report_table As New DataTable
        Dim sqlString As String
        Dim nDataSource As DataTable

        Dim sumList As New ArrayList

        Dim fileTypeDDL As New DropDownList
        Dim fileType As String = ""

        Dim wMsg, cwMsg As String

        Try
            If isPreview Then
                fileType = "PREVIEW"
            Else
                If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                    fileTypeDDL = DirectCast(hddiv.FindControl("rpt_save_as"), DropDownList)

                    If fileTypeDDL.SelectedValue.ToString <> "" Then
                        fileType = fileTypeDDL.SelectedItem.Value.ToString.ToUpper
                    Else
                        If Session("gLang") = "C" Then
                            uiFun.displayMsg(Me, "", "你必須選擇文件類型(另存為)來生成報告!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "You must select the Type of File(Save As) to generate Report!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    fileType = "EXCEL"
                End If
            End If

            If isColSelection Then
                report_table = gfBuildDataTableforGridView(datasource, gvrsList)
                dt_report_table = gfBuildDataTableforGridView(dt_datasource, gvrsList2)

                wMsg = "You must select Field to generate report!"
                cwMsg = "您必須選擇欄位!"
            Else
                report_table = datasource.Copy
                dt_report_table = dt_datasource.Copy

                wMsg = "No Column is available to be generated!"
                cwMsg = "沒有列可供產生!"
            End If

            If Not report_table Is Nothing And Not dt_report_table Is Nothing Then
                If dt_report_table.Rows.Count > 0 Then
                    report_table.Merge(dt_report_table)
                End If
            End If

            If report_table Is Nothing Then
                If Session("gLang") = "C" Then
                    uiFun.displayMsg(Me, "", cwMsg, Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", wMsg, Session("gLang"))
                End If

                Exit Sub
            ElseIf report_table.Rows.Count = 0 Then
                If Session("gLang") = "C" Then
                    uiFun.displayMsg(Me, "", cwMsg, Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", wMsg, Session("gLang"))
                End If

                Exit Sub
            End If

            Dim Sort_Col As String = "cold_display_seq ASC"
            Dim rptTable As New DataTable

            rptTable = report_table.Clone

            Dim foundRows As DataRow() = report_table.Select("", Sort_Col)

            For i As Integer = 0 To foundRows.Count - 1
                rptTable.ImportRow(foundRows(i))
            Next

            rptTable.AcceptChanges()

            report_table = rptTable

            rptTable.Dispose()

            sqlString = generateSQL(report_table, sumList)


            nDataSource = gDB.getDataTable(sqlString)

            generateSumColumn(nDataSource, sumList)

            RptU.RenderReport(menu_code, nDataSource, report_table, menu_code, fileType, table_name, Session("gLang"), rpt_title, Me.ClientScript)

        Catch ex As Exception
            Me.Controls.Clear()
            Response.Write(ex.Message)
        End Try
    End Sub

    Private Sub GenerateHtmlPreview(ByVal valTable As DataTable)
        Dim gv As New GridView
        Dim bField As BoundField
        Dim fieldLabel As String = ""

        If Session("gLang") = "C" Then
            fieldLabel = "cold_chi_label"
        Else
            fieldLabel = "cold_eng_label"
        End If

        For i As Integer = 0 To valTable.Rows.Count - 1
            bField = New BoundField

            bField.DataField = valTable.Rows(i).Item("table_name").ToString & "_" & valTable.Rows(i).Item("column_name").ToString

            bField.HeaderText = valTable.Rows(i).Item("fieldLabel").ToString.Trim

            bField.ItemStyle.CssClass = "GV"

            gv.Columns.Add(bField)
        Next

        gv.ID = menu_code & "_gvView"
        gv.Style.Add("width", "100%")
        gv.AllowPaging = False
        gv.Font.Name = "Arial"
        gv.Font.Overline = False
        gv.AutoGenerateColumns = False
        gv.PageSize = 20

        gv.BorderStyle = BorderStyle.None
        gv.BorderWidth = 1
        gv.CellPadding = 3
        gv.CaptionAlign = TableCaptionAlign.Top

        gv.RowStyle.BorderColor = Drawing.Color.Silver
        gv.RowStyle.BorderStyle = BorderStyle.Solid
        gv.RowStyle.BorderWidth = 1

        gv.HeaderStyle.BorderColor = Drawing.Color.White
        gv.HeaderStyle.BorderStyle = BorderStyle.Inset
        gv.HeaderStyle.BorderWidth = 1
        gv.HeaderStyle.Font.Name = "Arial"
        gv.HeaderStyle.Font.Bold = True
        gv.HeaderStyle.Font.Size = 9
        gv.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
        gv.HeaderStyle.VerticalAlign = VerticalAlign.Bottom
        gv.HeaderStyle.Wrap = True
        gv.HeaderStyle.CssClass = "GridViewFixedHeader"

        gv.FooterStyle.BackColor = Drawing.Color.White
        gv.SelectedRowStyle.Font.Bold = True
        gv.SelectedRowStyle.ForeColor = Drawing.Color.White

        gv.PagerStyle.Font.Size = 8
        gv.PagerStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#4C60B6")
        gv.PagerStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#006699")
        gv.PagerStyle.HorizontalAlign = HorizontalAlign.Left
        gv.PagerStyle.BorderColor = Drawing.Color.Transparent
        gv.PagerStyle.BorderWidth = 1
        gv.PagerStyle.Font.Bold = True
        gv.PagerStyle.Font.Name = "Arial"
        gv.PagerStyle.Font.Strikeout = False

        gv.EmptyDataText = "No Record Found."

        gv.Style.Add("width", "97%")
        'mPanel.Controls.Add(gv)


    End Sub

    Private Function GenerateCell(ByVal tblRow As DataRow) As TableCell
        Dim cell As New TableCell()

        Select Case tblRow.Item("srcl_edit_style").ToString.Trim
            Case "T"
                Dim tb As New TextBox()

                tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                tb.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")

                If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue _
                                        Where defKey.Contains("SESSION") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        tb.Text = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    valueQuery = From defKey In defValue _
                                        Where defKey.Contains("REQUEST") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        tb.Text = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then tb.Text = tblRow.Item("srcl_def_value_or_selected_seq").ToString
                End If


                cell.Controls.Add(tb)

            Case "D"
                If tblRow.Item("srcl_date_range_yn").ToString.Trim.ToUpper = "N" Then
                    Dim tb As New TextBox
                    Dim link As New HyperLink
                    Dim img As New Image

                    tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb.Width = 80
                    tb.MaxLength = 10
                    tb.Attributes("onkeypress") = "return maskDate(event);"

                    If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                        If IsNumeric(tblRow.Item("srcl_def_value_or_selected_seq").ToString) Then
                            tb.Text = Format(Now.AddDays(CDbl(tblRow.Item("srcl_def_value_or_selected_seq").ToString)), "yyyy/MM/dd")
                        End If
                    End If

                    link.Attributes("target") = "_self"
                    link.Attributes("onclick") = "javascript:pedirFecha(" & tb.ID & ",'D_date');"

                    link.Style.Add("cursor", "hand")

                    img.ImageUrl = "images/calendar.gif"
                    img.ID = tb.ID & "_img"
                    img.ImageAlign = ImageAlign.AbsMiddle
                    img.BorderWidth = 0
                    img.Style.Add("padding-left", "4px")

                    link.Controls.Add(img)

                    If NoAccess Then
                        link.Enabled = False
                    End If

                    cell.Controls.Add(tb)
                    cell.Controls.Add(link)
                Else
                    Dim tb1 As New TextBox
                    Dim tb2 As New TextBox
                    Dim link1 As New HyperLink
                    Dim link2 As New HyperLink
                    Dim img1 As New Image
                    Dim img2 As New Image

                    Dim lbl1 As New Label
                    Dim lbl2 As New Label

                    tb1.ID = "FR_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb1.Width = 80
                    tb1.MaxLength = 10
                    tb1.Attributes("onkeypress") = "return maskDate(event);"

                    If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                        Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

                        For i As Integer = 0 To defValue.Count - 1
                            If i <= 2 Then
                                If i = 0 Then
                                    tb1.Text = Format(Now.AddDays(CDbl(defValue(i).ToString)), "yyyy/MM/dd")
                                ElseIf i = 1 Then
                                    tb2.Text = Format(Now.AddDays(CDbl(defValue(i).ToString)), "yyyy/MM/dd")
                                End If
                            End If
                        Next
                    End If

                    link1.Attributes("target") = "_self"
                    link1.Attributes("onclick") = "javascript:pedirFecha(" & tb1.ID & ",'D_date');"

                    link1.Style.Add("cursor", "hand")

                    img1.ImageUrl = "images/calendar.gif"
                    img1.ID = tb1.ID & "_img"
                    img1.ImageAlign = ImageAlign.AbsMiddle
                    img1.BorderWidth = 0
                    img1.Style.Add("padding-left", "4px")

                    tb2.ID = "TO_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb2.Width = 80
                    tb2.MaxLength = 10
                    tb2.Attributes("onkeypress") = "return maskDate(event);"

                    link2.Attributes("target") = "_self"
                    link2.Attributes("onclick") = "javascript:pedirFecha(" & tb2.ID & ",'D_date');"
                    link2.Style.Add("cursor", "hand")

                    img2.ImageUrl = "images/calendar.gif"
                    img2.ID = tb2.ID & "_img"
                    img2.ImageAlign = ImageAlign.AbsMiddle
                    img2.BorderWidth = 0
                    img2.Style.Add("padding-left", "4px")

                    If Session("gLang") = "C" Then
                        lbl1.Text = "由 "
                        lbl2.Text = " 至 "
                    Else
                        lbl1.Text = "From "
                        lbl2.Text = " To "
                    End If

                    link1.Controls.Add(img1)
                    link2.Controls.Add(img2)

                    If NoAccess Then
                        link1.Enabled = False
                        link2.Enabled = False
                    End If

                    cell.Controls.Add(lbl1)
                    cell.Controls.Add(tb1)
                    cell.Controls.Add(link1)
                    cell.Controls.Add(lbl2)
                    cell.Controls.Add(tb2)
                    cell.Controls.Add(link2)
                End If

            Case "C"
                Dim cbl As New CheckBoxList
                cbl.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                cbl.CssClass = "NEWTEXT"

                If tblRow.Item("srcl_checkbox_rowcnt").ToString <> "" Then _
                    cbl.RepeatColumns = CInt(tblRow.Item("srcl_checkbox_rowcnt").ToString)

                cbl.RepeatDirection = RepeatDirection.Horizontal

                Dim li As ListItem

                Dim cbString As String = ""

                Dim cbTbl As New DataTable
                Dim gotSeq As Boolean = False

                If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                    cbString = "select colc_code as code, colc_eng_value as name, colc_chi_value as cname, colc_display_seq from " & cms_col_code & " where colc_tabcol = '" & _
                                            tblRow.Item("cold_tabcol").ToString.Trim & "' " & _
                                            "order by colc_display_seq"
                    gotSeq = True
                Else
                    cbString = tblRow.Item("srcl_sql").ToString.Trim
                End If

                cbTbl = gDB.getDataTable(cbString)

                Dim tempCheckedList As String() = Nothing

                Dim defArrayList As New ArrayList
                If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

                    Dim valueQuery = From defKey In defValue _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        defArrayList.Add(dVal)
                    Next
                End If

                For x As Integer = 0 To cbTbl.Rows.Count - 1
                    li = New ListItem

                    If Session("gLang") = "C" Then
                        li.Text = cbTbl.Rows(x).Item("cname").ToString
                    Else
                        li.Text = cbTbl.Rows(x).Item("name").ToString
                    End If

                    li.Value = cbTbl.Rows(x).Item("code").ToString

                    If gotSeq Then
                        If defArrayList.Contains(cbTbl.Rows(x).Item("colc_display_seq").ToString.Trim) Then
                            li.Selected = True
                        End If
                    Else
                        If defArrayList.Contains(x.ToString) Then
                            li.Selected = True
                        End If
                    End If

                    cbl.Items.Insert(x, li)
                Next

                cell.Controls.Add(cbl)

            Case "S"
                Dim ddl As New DropDownList
                ddl.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                    uiFun.load_dropdownBy_ColCode(ddl, tblRow.Item("cold_tabcol").ToString.Trim, Session("gLang"))
                Else
                    uiFun.load_dropdown(ddl, tblRow.Item("srcl_sql").ToString.Trim, , , , Session("gSelectLabel"))
                End If

                If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue _
                                        Where defKey.Contains("SESSION") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        ddl.SelectedValue = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    valueQuery = From defKey In defValue _
                                        Where defKey.Contains("REQUEST") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        ddl.SelectedValue = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then ddl.SelectedValue = tblRow.Item("srcl_def_value_or_selected_seq").ToString
                End If

                cell.Controls.Add(ddl)

            Case "L"
                Dim lbl As New Label

                lbl.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                lbl.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")

                If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue _
                                        Where defKey.Contains("SESSION") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        lbl.Text = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    valueQuery = From defKey In defValue _
                                        Where defKey.Contains("REQUEST") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        lbl.Text = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then lbl.Text = tblRow.Item("srcl_def_value_or_selected_seq").ToString
                End If

                cell.Controls.Add(lbl)

            Case "H"
                Dim hf As New HiddenField

                hf.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                If tblRow.Item("srcl_def_value_or_selected_seq").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("srcl_def_value_or_selected_seq").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue _
                                        Where defKey.Contains("SESSION") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        hf.Value = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    valueQuery = From defKey In defValue _
                                        Where defKey.Contains("REQUEST") _
                                        Select defKey

                    For Each dVal As String In valueQuery
                        hf.Value = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then hf.Value = tblRow.Item("srcl_def_value_or_selected_seq").ToString
                End If

                rptform.Controls.Add(hf)

            Case Else
                Dim tb As New TextBox

                tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                tb.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")


                cell.Controls.Add(tb)
        End Select

        GenerateCell = cell

    End Function

    Private Sub GenerateObjectSQL(ByRef SCString As String, ByVal tblRow As DataRow, ByVal srcl_edit_style As String, ByVal oper As String, _
                          ByVal ctrlName As String, ByVal FieldName As String, ByVal symbol As String, Optional ByVal org_edit_style As String = "")

        Dim item_srch_sql As String = ""

        Select Case srcl_edit_style
            Case "T"
                Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                Select Case oper.ToUpper
                    Case "="
                        item_srch_sql = "UPPER(" & FieldName & ") = N'" & gU.dbEncode(tmpTextbox.Text) & "' "
                    Case "LIKE"
                        item_srch_sql = FieldName & " LIKE N'%" & gU.dbEncode(tmpTextbox.Text) & "%' "
                    Case Else
                        Exit Sub
                End Select

                If tmpTextbox.Text <> "" Then
                    SCString = APPSQL(SCString, item_srch_sql, symbol)
                End If

            Case "D"
                If tblRow.Item("srcl_date_range_yn").ToString.Trim.ToUpper = "N" Then
                    Dim tmpDateTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                    If tmpDateTextbox.Text <> "" Then
                        SCString = APPSQL(SCString, FieldName & " >= N'" & gU.dbEncode(tmpDateTextbox.Text) & "' ", symbol)
                        SCString = APPSQL(SCString, FieldName & " < DATEADD(DAY, 1, N'" & gU.dbEncode(tmpDateTextbox.Text) & "') ", symbol)
                    End If
                Else
                    Dim tmpFromTextbox As TextBox = DirectCast(hddiv.FindControl("FR_" & ctrlName), TextBox)
                    Dim tmpToTextbox As TextBox = DirectCast(hddiv.FindControl("TO_" & ctrlName), TextBox)

                    If tmpFromTextbox.Text <> "" Then
                        SCString = APPSQL(SCString, FieldName & " >= N'" & gU.dbEncode(tmpFromTextbox.Text) & "' ", symbol)
                    End If

                    If tmpToTextbox.Text <> "" Then
                        SCString = APPSQL(SCString, FieldName & " < DATEADD(DAY, 1, N'" & gU.dbEncode(tmpToTextbox.Text) & "') ", symbol)
                    End If
                End If

            Case "C"
                Dim tmpCBValue As String = ""

                Dim tmpCheckBoxList As New CheckBoxList
                tmpCheckBoxList = DirectCast(hddiv.FindControl(ctrlName), CheckBoxList)

                If tmpCheckBoxList.SelectedIndex <> -1 Then
                    SCString = APPSQL(SCString, "UPPER(" & FieldName & ") IN (" & uiFun.getValueFromCB(tmpCheckBoxList) & ") ", symbol)

                    Dim tmpstr As String = ""
                    Dim tS As String = ""
                    Dim addSQL As String = ""

                    For i As Integer = 0 To tmpCheckBoxList.Items.Count - 1
                        If tmpCheckBoxList.Items(i).Selected = True Then
                            If tmpstr <> "" Then tmpstr = tmpstr & ","
                            tmpstr = tmpstr & tmpCheckBoxList.Items(i).Value

                            If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                                tS = "select colc_add_where_sql from " & cms_col_code & " where " & _
                                        "UPPER(colc_tabcol) = '" & tblRow.Item("cold_tabcol").ToString.Trim.ToUpper & "' and " & _
                                        "UPPER(colc_code) = '" & tmpCheckBoxList.Items(i).Value.ToUpper.ToString & "' "
                                addSQL = dbFun.getValueFromSQL(tS).ToUpper.ToString

                                If addSQL <> "" Then
                                    SCString = SCString & " " & addSQL
                                End If
                            End If
                        End If
                    Next
                End If

            Case "S"
                Dim tmpDDLValue As String = ""

                Dim tmpDropDownList As New DropDownList
                tmpDropDownList = DirectCast(hddiv.FindControl(ctrlName), DropDownList)

                If tmpDropDownList.SelectedValue.ToString <> "" Then
                    SCString = APPSQL(SCString, "UPPER(" & FieldName & ") = N'" & gU.dbEncode(tmpDropDownList.SelectedItem.Value) & "' ", symbol)
                End If

            Case "L"
                Dim tmpLabel As Label = TryCast(hddiv.FindControl(ctrlName), Label)

                If tmpLabel IsNot Nothing Then
                    Select Case oper.ToUpper
                        Case "="
                            item_srch_sql = "UPPER(" & FieldName & ") = N'" & gU.dbEncode(tmpLabel.Text.Trim) & "' "
                        Case "LIKE"
                            item_srch_sql = FieldName & " LIKE N'%" & gU.dbEncode(tmpLabel.Text.Trim) & "%' "
                        Case Else
                            Exit Sub
                    End Select

                    If tmpLabel.Text.Trim <> "" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                Else
                    GenerateObjectSQL(SCString, tblRow, org_edit_style, oper, ctrlName, FieldName, symbol)
                End If

            Case "H"
                Dim tmpHidden As HiddenField = TryCast(rptform.FindControl(ctrlName), HiddenField)

                If tmpHidden IsNot Nothing Then
                    Session("SEARCH_SESSION_PAGE_" & tmpHidden.ID) = tmpHidden.Value

                    Select Case oper.ToUpper
                        Case "="
                            item_srch_sql = "UPPER(" & FieldName & ") = N'" & gU.dbEncode(tmpHidden.Value) & "' "
                        Case "LIKE"
                            item_srch_sql = FieldName & " LIKE N'%" & gU.dbEncode(tmpHidden.Value) & "%' "
                        Case Else
                            Exit Sub
                    End Select

                    If tmpHidden.Value <> "" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If
                Else
                    GenerateObjectSQL(SCString, tblRow, org_edit_style, oper, ctrlName, FieldName, symbol)
                End If

            Case Else
                Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                Select Case oper.ToUpper
                    Case "="
                        item_srch_sql = "UPPER(" & FieldName & ") = N'" & gU.dbEncode(tmpTextbox.Text) & "' "
                    Case "LIKE"
                        item_srch_sql = FieldName & " LIKE N'%" & gU.dbEncode(tmpTextbox.Text) & "%' "
                    Case Else
                        Exit Sub
                End Select

                If tmpTextbox.Text <> "" Then
                    SCString = APPSQL(SCString, item_srch_sql, symbol)
                End If
        End Select
    End Sub
End Class
