Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System
Imports System.Reflection
Imports System.Configuration
Imports System.Collections
Imports System.Xml
Imports System.Data.SqlClient
Imports ExcelLibrary.SpreadSheet

Partial Class cms_search
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private RptU As New ReportUtils
    Private ar As New AccessRightUtils
    Private appCon As New AppConfig
    Private dbFun As New DBfunc

    Private menu_code As String = ""
    Private srch_table As New DataTable
    Private col_table As New DataTable
    Private master_list_table As New DataTable
    Private list_table As New DataTable
    Private pForm As String = ""
    Private pFunc As String = ""
    Private pItemArray As String()

    Private remain_status As Boolean = False
    Private searched As Boolean = False

    Private mc_LList As New ArrayList

    Private newDataArray As String()

    Private NoAccess As Boolean = False

    Public Const SELECTED_CB_INDEX As String = "SelectedCBIndex"

    Private ReadOnly Property SelectedCBIndex() As List(Of Int32)
        Get
            If ViewState(SELECTED_CB_INDEX) Is Nothing Then
                ViewState(SELECTED_CB_INDEX) = New List(Of Int32)()
            End If

            Return DirectCast(ViewState(SELECTED_CB_INDEX), List(Of Int32))
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim pAr As String = ""
        Dim pItemList As String = ""
        Dim srch_sql As String = ""

        Dim showPrintTable As Boolean = False

        Dim showlistNow As Boolean = False



        Try
            'If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            '    Session.Remove("PAGE_SESSION_MENU_CODE")
            '    ar.Force_PageEndCtrlClear(Me, False)
            '    Me.Visible = False
            '    Exit Sub
            'End If
            ar.checkLoginStatus()

            If Not IsPostBack Then
                For i As Integer = 0 To (Request.Params.AllKeys.Count - 1)
                    If Request.Params.Keys(i).Trim.ToLower <> "asp.net_sessionid" Then
                        ViewState(Request.Params.Keys(i)) = Request.Params(i)
                    Else
                        Exit For
                    End If
                Next
                'ViewState("menu_code") = Request("menu_code")

                Dim slHD As New HiddenField
                slHD.ID = "sl"
                slHD.Value = HttpContext.Current.Request("sl")
                srchForm.Controls.Add(slHD)

                Dim mWHHD As New HiddenField
                mWHHD.ID = "cmWH"
                mWHHD.Value = HttpContext.Current.Request("cmWH")
                srchForm.Controls.Add(mWHHD)

            End If

            menu_code = ViewState("menu_code")

            'ar = New AccessRightUtils(menu_code, Session("usr_id"), Me)

            'NoAccess = ar.hideForm(Me)

            If menu_code = "" Then
                menu_code = Session("PAGE_SESSION_MENU_CODE")
            End If

            If NoAccess Then
                Session.Remove("PAGE_SESSION_MENU_CODE")
                Me.Visible = False
                Exit Sub
            End If

            If Session("SEARCH_SESSION_PAGE_SEARCHED") IsNot Nothing Then
                If Session("SEARCH_SESSION_PAGE_SEARCHED") = "Y" Then
                    searched = True
                End If
            End If

            If Not IsPostBack Then
                srch_table = Nothing
                col_table = Nothing
                list_table = Nothing
                master_list_table = Nothing

                'ViewState("pForm") = Request("pForm")
                'ViewState("pFunc") = Request("pFunc")
                'ViewState("pItemList") = Request("pItemList")
                ViewState("pfdt") = Request("fdt")
                ViewState("pfdn") = Request("fdn")
                'ViewState("pAr") = Request("ar")

                ViewState("pSTORER_CODE") = Request("sc")
                ViewState("pItem_code") = Request("ic")
                ViewState("pPack_key") = Request("pKey")
                ViewState("noresetbtn") = Request("noreset")
                ViewState("pILOC_WH") = Request("wh")
                ViewState("pWH_MAIN_WH") = Request("mwh")

                ViewState("showlistNow") = Request("sl")
            End If

            pForm = ViewState("pForm")
            pFunc = ViewState("pFunc")
            pItemList = ViewState("pItemList")
            pAr = ViewState("pAr")

            pItemArray = Split(pItemList, ", ")
            newDataArray = Split(pItemList, ", ")

            srch_table = ViewState("g_srch_table")
            col_table = ViewState("g_col_table")
            list_table = ViewState("g_list_col_table")
            master_list_table = ViewState("g_master_list_col_table")

            showPrintTable = ViewState("isShowPrintTable")

            If ViewState("showlistNow") IsNot Nothing Then
                showlistNow = ViewState("showlistNow")
            End If

            'mc_LList = ViewState("g_mc_linkList")

            If menu_code = "" Then
                Session.Remove("PAGE_SESSION_MENU_CODE")

                If Not NoAccess Then
                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: menu code is missing.');", True)
                End If

                Exit Sub
            End If

            If srch_table Is Nothing Then
                srch_sql = "select wms_search.*, " &
                                    "wms_function.fun_eng_name, wms_function.fun_chi_name, " &
                                    "wms_function.fun_rpt_eng_name, wms_function.fun_rpt_chi_name " &
                                " from wms_search, wms_function where " &
                                    "wms_search.fun_code = wms_function.fun_code and " &
                                    "wms_search.fun_code = '" & menu_code & "'"

                srch_table = gDB.getDataTable(srch_sql)

                ViewState("g_srch_table") = srch_table

                If srch_table.Rows.Count > 0 Then
                    If srch_table.Rows(0).Item("SRCH_TYPE").ToString <> "LOOKUP" And srch_table.Rows(0).Item("SRCH_TYPE").ToString <> "REPORT" And ViewState("noresetbtn") <> "Y" Then
                        Session.Remove("PAGE_SESSION_MENU_CODE")
                        Session("PAGE_SESSION_MENU_CODE") = menu_code

                        If Not Session("SEARCH_SESSION_PAGE_REMAIN_STATUS") Is Nothing Then
                            If Session("SEARCH_SESSION_PAGE_REMAIN_STATUS") Then
                                If Not Session("SEARCH_SESSION_PAGE_STATUS_MENU_CODE") Is Nothing Then
                                    If Session("SEARCH_SESSION_PAGE_STATUS_MENU_CODE") = menu_code Then
                                        remain_status = True
                                    Else
                                        appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")
                                    End If
                                End If
                            Else
                                appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")
                            End If
                        End If
                    Else
                        Session.Remove("LR_SESSION_MENU_CODE")
                        Session("LR_SESSION_MENU_CODE") = menu_code
                    End If

                    If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                        GenerateSearchTable()
                        GenerateGV()
                        GenerateMasterGV()
                    Else
                        GenerateSearchTable()
                        GenerateGV()
                    End If

                    If srch_table.Rows(0).Item("SRCH_START_WITH_SEARCH_YN").ToString.ToUpper.Trim = "Y" Then
                        showlistNow = True
                    End If
                Else
                    If Not NoAccess Then
                        System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: wms_search define data is missing.');", True)
                    End If

                    Exit Sub
                End If
            Else
                If srch_table.Rows.Count > 0 Then
                    If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                        GenerateSearchTable()
                        GenerateGV()
                        GenerateMasterGV()

                        Dim gv As GridView = DirectCast(mdiv.FindControl(menu_code & "_MasterGridView"), GridView)

                        If gv.Rows.Count > 0 Then
                            ClientScript.RegisterStartupScript(Me.GetType(), "CreateGridHeader", "<script>CreateGridHeader('mGVDiv', '" & gv.ClientID & "', 'mGVHDDiv');</script>")
                        End If
                    Else
                        GenerateSearchTable()
                        GenerateGV()
                    End If

                    If srch_table.Rows(0).Item("SRCH_START_WITH_SEARCH_YN").ToString.ToUpper.Trim = "Y" Then
                        showlistNow = True
                    End If
                End If
            End If

            If Not pAr Is Nothing Then
                If pAr.Trim.ToUpper = "Y" Then

                    Dim sBtn As Button = DirectCast(srchForm.FindControl("btnSearch"), Button)
                    sBtn.Visible = False

                    showlistNow = True
                End If
            End If

            If showPrintTable Then
                If TryCast(dtdiv.FindControl("printTable"), Table) IsNot Nothing Then
                    DirectCast(dtdiv.FindControl("printTable"), Table).Visible = True
                End If
            End If

            If ((remain_status And searched) Or showlistNow) AndAlso srch_table.Rows(0).Item("SRCH_DIRECTLY_REDIRECT_YN").ToString.Trim <> "Y" Then
                Call BindGV()
            End If


            If Not IsPostBack Then
                Page.Form.Focus()
            End If

            If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
                Dim checkedCount As New HiddenField
                checkedCount.ID = "chkCount"
                checkedCount.Value = 0
                srchForm.Controls.Add(checkedCount)
                Dim cbJs As String = ""

                cbJs += "<script language='JavaScript'>" & vbNewLine
                cbJs += "function ChgHDState(obj) {" & vbNewLine
                cbJs += "   if (chkArray != null) {" & vbNewLine
                cbJs += "       var rc = document.getElementById('RowsCount');" & vbNewLine
                cbJs += "       var cc = document.getElementById('chkCount');" & vbNewLine
                cbJs += "       if (obj.checked) {" & vbNewLine
                cbJs += "               cc.value = parseFloat(cc.value) + 1;" & vbNewLine
                cbJs += "       } else {" & vbNewLine
                cbJs += "               cc.value = parseFloat(cc.value) - 1;" & vbNewLine
                cbJs += "       } " & vbNewLine
                cbJs += "       if (rc != null && cc != null) {" & vbNewLine
                cbJs += "           if (rc.value == cc.value) {" & vbNewLine
                cbJs += "               //ChgChkState(chkArray[0], true);" & vbNewLine
                cbJs += "               return;" & vbNewLine
                cbJs += "           } else {" & vbNewLine
                cbJs += "               //ChgChkState(chkArray[0], false);" & vbNewLine
                cbJs += "               return;" & vbNewLine
                cbJs += "           }" & vbNewLine
                cbJs += "       }" & vbNewLine
                cbJs += "   }" & vbNewLine


                cbJs += "}" & vbNewLine
                cbJs += "</script>"

                ClientScript.RegisterStartupScript(Me.GetType, "chgcbHD", cbJs)
            End If

        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(srch_sql.ToUpper) & "\n\r\n\r" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Private Sub GenerateSearchTable()
        Dim tblSQL As String = "select * from wms_search_col where fun_code = '" & menu_code & "' and srcl_search_seq <> 0 order by srcl_search_seq"
        Dim tbl As New DataTable
        Dim field As String
        Dim page_title As String
        Try
            tbl = gDB.getDataTable(tblSQL)

            ViewState("g_col_table") = tbl
            col_table = ViewState("g_col_table")

            If Session("gLang") = "C" Then
                field = "srcl_chi_label"
                page_title = "fun_chi_name"
            Else
                field = "srcl_label"
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

            tit_cell1.CssClass = "TITLE"
            tit_cell1.Style.Add("width", "100%")

            Dim tit_Label As New Label()
            tit_Label.ID = "lblTitle_N"
            tit_Label.Text = srch_table.Rows(0).Item(page_title).ToString
            tit_Label.Font.Bold = True
            tit_cell1.Controls.Add(tit_Label)

            ViewState("rpt_title") = srch_table.Rows(0).Item(page_title).ToString

            tit_row.Cells.Add(tit_cell1)

            If srch_table.Rows(0).Item("srch_show_new_button_yn").ToString.Trim.ToUpper = "Y" Then
                Dim tit_cell2 As New TableCell()

                tit_cell2.Wrap = False
                tit_cell2.CssClass = "TITLE"
                tit_cell2.Style.Add("width", "100%")

                Dim newBtn As New Button
                newBtn.ID = menu_code & "_NewBtn"
                newBtn.CssClass = "all_button"
                newBtn.Text = "New"

                Dim sm As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                sm.RegisterPostBackControl(newBtn)

                AddHandler newBtn.Click, AddressOf new_redirect

                tit_cell2.Controls.Add(newBtn)
                tit_row.Cells.Add(tit_cell2)
            Else
                tit_cell1.ColumnSpan = 2
            End If

            tit_table.Rows.Add(tit_row)

            headercell.Controls.Add(tit_table)

            headerrow.Cells.Add(headercell)
            table.Rows.Add(headerrow)

            Dim TotlColCount As Integer = 0
            Dim colCount As Integer = 0
            Dim rowIdx As Integer = 0

            Dim row As New TableRow
            'row.ID = "RW1"

            Dim rowsArray As New ArrayList

            For i As Integer = 0 To tbl.Rows.Count - 1
                'Dim row As New TableRow()
                Dim tCell As New TableCell()
                Dim tLabel As New Label()

                If colCount > TotlColCount Then TotlColCount = colCount

                If tbl.Rows(i).Item("srcl_same_row_yn").ToString.Trim.ToUpper = "Y" Then
                    If rowsArray.Count > 1 Then
                        Dim tmprow As TableRow = DirectCast(rowsArray(rowIdx - 1), TableRow)
                        If tmprow.Cells.Count < colCount Then
                            Dim spanNo As Integer = 0

                            spanNo = colCount - tmprow.Cells.Count
                            tmprow.Cells(tmprow.Cells.Count - 1).ColumnSpan = spanNo
                        End If
                    End If
                Else
                    If i <> 0 Then
                        If (tbl.Rows(i).Item("srcl_edit_style").ToString.Trim <> "H" AndAlso tbl.Rows(i).Item("srcl_edit_style").ToString.Trim <> "B") OrElse
                            (tbl.Rows(i).Item("srcl_edit_style").ToString.Trim = "H" AndAlso tbl.Rows(i).Item("srcl_app_custom_ctrl_links").ToString.Trim <> "") Then
                            table.Rows.Add(row)

                            rowsArray.Add(row)

                            colCount = 0

                            rowIdx += 1

                            row = New TableRow
                            row.ID = "RW" & rowIdx
                        End If
                    End If
                End If

                tLabel.ID = "lbl" & tbl.Rows(i).Item("cold_tabcol").ToString.Replace(".", "_")
                tLabel.Text = tbl.Rows(i).Item(field).ToString & ":"

                tLabel.Font.Size = 10
                tCell.Controls.Add(tLabel)
                tCell.CssClass = "LabelTD"
                tCell.Style.Add("width", "20%")
                tCell.Wrap = False

                row.Cells.Add(tCell)

                If tbl.Rows(i).Item("SRCL_ISREADONLY_YN").ToString.Trim = "Y" Then
                    If tbl.Rows(i).Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.Trim = "" Then
                        Dim itmCell As TableCell = GenerateCell(tbl.Rows(i), i)
                        If itmCell IsNot Nothing Then row.Cells.Add(itmCell)
                    Else
                        Dim objValue As String = ""

                        Dim defValue As String() = tbl.Rows(i).Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                        Dim elseCase As Boolean = True

                        Dim valueQuery = From defKey In defValue
                                         Where defKey.ToUpper.Contains("SESSION")
                                         Select defKey

                        For Each dVal As String In valueQuery
                            objValue = Session(dVal.ToUpper.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))

                            If objValue IsNot Nothing Then
                                If objValue <> "" Then
                                    ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                                End If

                                elseCase = False
                                Exit For
                            End If
                        Next

                        If elseCase Then
                            valueQuery = From defKey In defValue
                                         Where defKey.ToUpper.Contains("REQUEST")
                                         Select defKey

                            For Each dVal As String In valueQuery
                                objValue = Request(dVal.ToUpper.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))

                                If objValue IsNot Nothing Then
                                    elseCase = False

                                    If objValue <> "" Then
                                        ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                                    End If

                                    Exit For
                                End If
                            Next
                        End If

                        If elseCase Then
                            valueQuery = From defKey In defValue
                                         Select defKey

                            For Each dVal As String In valueQuery
                                objValue = ViewState(dVal)

                                If objValue IsNot Nothing Then
                                    elseCase = False

                                    If objValue <> "" Then
                                        ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                                    End If
                                End If

                                Exit For
                            Next
                        End If

                        If elseCase Then objValue = tbl.Rows(i).Item("SRCL_DEF_VALUE_SEL_SEQ").ToString

                        If objValue Is Nothing Then objValue = ""

                        Dim cell As New TableCell()

                        Select Case tbl.Rows(i).Item("srcl_edit_style").ToString.Trim
                            Case "T", "L"
                                Dim lbl As New Label

                                lbl.ID = tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                                lbl.Width = gU.decodeEmptyCInt(tbl.Rows(i).Item("srcl_col_width").ToString, "120")

                                If objValue.Trim <> "" Then
                                    lbl.Text = objValue.Trim
                                    ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue.Trim
                                Else
                                    lbl.Text = ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE")
                                End If

                                cell.Controls.Add(lbl)

                            Case "S"
                                Dim ddl As New DropDownList

                                If tbl.Rows(i).Item("srcl_sql").ToString.Trim = "" Then
                                    uiFun.load_dropdownBy_ColCode(ddl, tbl.Rows(i).Item("cold_tabcol").ToString.Trim, Session("gLang"))
                                Else
                                    Dim app_symbol As String = tbl.Rows(i).Item("srcl_app_sql").ToString
                                    Dim ddlSQL As String = tbl.Rows(i).Item("srcl_sql").ToString.Trim

                                    ddlSQL += " " & APP_SESSION_SQL(tbl.Rows(i).Item("srcl_sessioncol").ToString, tbl.Rows(i).Item("srcl_sessionval").ToString, app_symbol)
                                    ddlSQL += " " & APP_REQUEST_SQL(tbl.Rows(i).Item("srcl_requestcol").ToString, tbl.Rows(i).Item("srcl_requestval").ToString, app_symbol)

                                    If gU.decodeNullOrEmpty(tbl.Rows(i).Item("srcl_add_sql").ToString, "").Trim <> "" Then
                                        ddlSQL += " " & tbl.Rows(i).Item("srcl_add_sql").ToString.Trim.ToUpper
                                    End If

                                    If Not ddlSQL.ToUpper.Contains("ORDER BY") Then
                                        ddlSQL += " ORDER BY 2"
                                    End If

                                    uiFun.load_dropdown(ddl, ddlSQL.Trim, , , , Session("gSelectLabel"))
                                End If

                                Dim lblHD As New HiddenField
                                Dim lbl As New Label

                                lblHD.ID = tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                                lbl.ID = "dispLbl_" & tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                                If objValue <> "" Then
                                    ddl.SelectedValue = objValue
                                    ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE") = objValue
                                Else
                                    ddl.SelectedValue = ViewState(tbl.Rows(i).Item("cold_tabcol").ToString.Trim.Replace(".", "_") & "_VIEWSTATE")
                                End If

                                lblHD.Value = ddl.SelectedValue
                                lbl.Text = ddl.SelectedItem.Text.Trim

                                If ddl.SelectedValue = "" Then
                                    ddl.ID = lblHD.ID

                                    If remain_status Then
                                        ddl.SelectedValue = Session("SEARCH_SESSION_PAGE_" & ddl.ID)
                                    End If

                                    lblHD.Dispose()
                                    lbl.Dispose()

                                    cell.Controls.Add(ddl)
                                Else
                                    srchForm.Controls.Add(lblHD)
                                    cell.Controls.Add(lbl)
                                End If
                        End Select

                        row.Cells.Add(cell)
                    End If
                Else
                    Dim itmCell As TableCell = GenerateCell(tbl.Rows(i), i)
                    If itmCell IsNot Nothing Then
                        row.Cells.Add(itmCell)
                    End If
                End If

                'If tbl.Rows(i).Item("srcl_edit_style").ToString.Trim <> "H" AndAlso tbl.Rows(i).Item("srcl_edit_style").ToString.Trim <> "B" Then
                '    table.Rows.Add(row)
                '    colCount += 2
                'ElseIf tbl.Rows(i).Item("srcl_edit_style").ToString.Trim = "H" AndAlso tbl.Rows(i).Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                '    table.Rows.Add(row)
                '    colCount += 2
                'End If

                colCount += 2
            Next

            rowsArray.Add(row)

            If colCount > TotlColCount Then TotlColCount = colCount

            For r As Integer = 0 To rowsArray.Count - 1
                Dim spanNo As Integer = 0

                If DirectCast(rowsArray(r), TableRow).Cells.Count < TotlColCount Then
                    spanNo = TotlColCount - DirectCast(rowsArray(r), TableRow).Cells.Count
                    DirectCast(rowsArray(r), TableRow).Cells(DirectCast(rowsArray(r), TableRow).Cells.Count - 1).ColumnSpan = spanNo + 1
                End If
                table.Rows.Add(rowsArray(r))
            Next

            'If TotlColCount = 0 Then TotlColCount = 2

            headercell.ColumnSpan = TotlColCount

            Dim buttonrow As New TableRow()
            Dim buttoncell As New TableCell()
            buttoncell.ColumnSpan = TotlColCount

            Dim btnTale As New Table
            btnTale.BorderWidth = 0
            btnTale.CellPadding = 1
            btnTale.CellSpacing = 0
            btnTale.Style.Add("width", "100%")
            btnTale.Style.Add("align", "center")

            btnTale.ID = "btnTbl"

            Dim btnrow As New TableRow()
            Dim btnCell As New TableCell()
            Dim btn As New Button

            btnCell.CssClass = "TITLE"
            btnCell.Style.Add("width", "100%")

            btnrow.VerticalAlign = VerticalAlign.Top

            btnTale.Rows.Add(btnrow)

            buttoncell.Controls.Add(btnTale)

            buttonrow.Cells.Add(buttoncell)

            table.Rows.Add(buttonrow)

            If srch_table.Rows(0).Item("SRCH_TYPE").ToString = "LOOKUP" And
                    srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" Then

                Dim searchcell As New TableCell()

                searchcell.VerticalAlign = VerticalAlign.Middle
                searchcell.Wrap = False
                searchcell.CssClass = "TITLE"
                searchcell.Style.Add("width", "100%")

                btn.ID = "btnSearch"
                btn.CssClass = "all_button"

                If Session("gLang") = "E" Then
                    btn.Text = "Search"
                ElseIf Session("gLang") = "C" Then
                    btn.Text = "搜尋"
                End If

                btn.Width = 63

                'srchForm.DefaultButton = "btnSearch"
                srchForm.DefaultButton = btn.UniqueID

                If srch_table.Rows(0).Item("SRCH_DIRECTLY_REDIRECT_YN").ToString.ToUpper.Trim = "Y" Then
                    Dim sm As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                    sm.RegisterPostBackControl(btn)
                End If

                AddHandler btn.Click, AddressOf click_search

                searchcell.Controls.Add(btn)

                btnrow.Cells.Add(searchcell)

                Dim btnSub As New Button
                btnSub.ID = "cbSubmit"
                btnSub.CssClass = "all_button"
                btnSub.Width = 50

                If Session("gLang") = "E" Then
                    btnSub.Text = "OK"
                ElseIf Session("gLang") = "C" Then
                    btnSub.Text = "選擇"
                End If

                AddHandler btnSub.Click, AddressOf cbSubmit_Click

                btnCell.Controls.Add(btnSub)
                btnrow.Cells.Add(btnCell)

                srch_table.Rows(0).Item("srch_show_sort_yn") = "N"
            Else
                btn.ID = "btnSearch"
                btn.CssClass = "all_button"

                If Session("gLang") = "E" Then
                    btn.Text = "Search"
                ElseIf Session("gLang") = "C" Then
                    btn.Text = "搜尋"
                End If

                btn.Width = 63

                'srchForm.DefaultButton = "btnSearch"
                srchForm.DefaultButton = btn.UniqueID

                If srch_table.Rows(0).Item("SRCH_DIRECTLY_REDIRECT_YN").ToString.ToUpper.Trim = "Y" Then
                    Dim sm As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                    sm.RegisterPostBackControl(btn)
                End If

                AddHandler btn.Click, AddressOf click_search

                btnCell.Controls.Add(btn)

                If ViewState("noresetbtn") <> "Y" And srch_table.Rows(0).Item("SRCH_TYPE").ToString <> "LOOKUP" And srch_table.Rows(0).Item("SRCH_TYPE").ToString <> "REPORT" Then
                    Dim btn1 As New Button

                    btn1.ID = "btnReset"
                    btn1.CssClass = "all_button"

                    If Session("gLang") = "E" Then
                        btn1.Text = "Reset"
                    ElseIf Session("gLang") = "C" Then
                        btn1.Text = "重置"
                    End If

                    btn1.Width = 63

                    AddHandler btn1.Click, AddressOf click_reset

                    'btnCell.Controls.Add(New LiteralControl("&nbsp;"))
                    btnCell.Controls.Add(btn1)
                End If

                btnrow.Cells.Add(btnCell)
            End If

            If gU.decodeNullOrEmpty(srch_table.Rows(0).Item("SRCH_SHOW_SORT_YN").ToString, "Y").Trim.ToUpper = "Y" Then
                Dim sortlbl_cell As New TableCell()
                Dim sort_cell As New TableCell()

                sortlbl_cell.VerticalAlign = VerticalAlign.Middle
                sortlbl_cell.Wrap = False
                sortlbl_cell.CssClass = "TITLE"
                sortlbl_cell.Style.Add("width", "5%")

                sort_cell.VerticalAlign = VerticalAlign.Middle
                sort_cell.Wrap = False
                sort_cell.CssClass = "TITLE"
                sort_cell.Style.Add("width", "100%")

                Dim sort_label As New Label
                sort_label.ID = "lblSort"
                sort_label.Height = 18
                sort_label.Font.Size = 10
                sort_label.Font.Bold = True

                If Session("gLang") = "E" Then
                    sort_label.Text = "Sorted By:&nbsp;"
                ElseIf Session("gLang") = "C" Then
                    sort_label.Text = "排序方式:&nbsp;"
                End If

                sortlbl_cell.Controls.Add(sort_label)
                btnrow.Cells.Add(sortlbl_cell)

                Dim sortDDL As New DropDownList
                Dim sortDDL2 As New DropDownList
                Dim sortDDL3 As New DropDownList

                Dim ddlStr As String
                sortDDL.ID = menu_code & "_orderlist"
                sortDDL.Font.Size = 9

                sortDDL2.ID = menu_code & "_orderlist2"
                sortDDL2.Font.Size = 9

                sortDDL3.ID = menu_code & "_orderlist3"
                sortDDL3.Font.Size = 9

                ddlStr = "select cold_tabcol as value, srcl_label, srcl_chi_label from wms_search_col where fun_code = '" & menu_code & "' and srcl_list_seq <> 0 and ISNULL(srcl_add_to_sort_yn,'Y') <> 'N' order by srcl_list_seq "

                uiFun.load_dropdown(sortDDL, ddlStr, "value", field, "", Session("gSelectLabel"))
                'uiFun.load_dropdown(sortDDL2, ddlStr, "value", field, "", Session("gSelectLabel"))
                'uiFun.load_dropdown(sortDDL3, ddlStr, "value", field, "", Session("gSelectLabel"))

                For Each item As ListItem In sortDDL.Items
                    If item.Value.Contains(".") Then
                        If item.Value.Split(".").Count > 0 Then
                            'item.Value = item.Value.Split(".")(1)
                            item.Value = item.Value.Replace(".", "@")
                        End If
                    End If
                Next

                'For Each item As ListItem In sortDDL2.Items
                '    If item.Value.Contains(".") Then
                '        If item.Value.Split(".").Count > 0 Then
                '            item.Value = item.Value.Split(".")(1)
                '        End If
                '    End If
                'Next

                'For Each item As ListItem In sortDDL3.Items
                '    If item.Value.Contains(".") Then
                '        If item.Value.Split(".").Count > 0 Then
                '            item.Value = item.Value.Split(".")(1)
                '        End If
                '    End If
                'Next

                If remain_status Then
                    sortDDL.SelectedValue = Session("SEARCH_SESSION_PAGE_" & sortDDL.ID)
                End If


                sort_cell.Controls.Add(sortDDL)
                'sort_cell.Controls.Add(sortDDL2)
                'sort_cell.Controls.Add(sortDDL3)

                btnrow.Cells.Add(sort_cell)

            Else
                If srch_table.Rows(0).Item("SRCH_TYPE").ToString <> "LOOKUP" And
                    srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper <> "Y" Then

                    btnCell.ColumnSpan = 2
                End If

            End If

            Dim printTbl_ADDED As Boolean = False
            Dim printTbl As New Table
            Dim printRow As New TableRow
            Dim printCell As New TableCell

            printTbl.ID = "printTable"
            printTbl.CellPadding = 0
            printTbl.CellSpacing = 0
            printTbl.Style.Add("width", "100%")
            printTbl.Style.Add("margin-bottom", "1px")
            printCell.VerticalAlign = VerticalAlign.Middle
            printCell.HorizontalAlign = HorizontalAlign.Right
            printCell.Style.Add("height", "25px")
            printCell.Wrap = False
            printCell.CssClass = "DtlLabel"


            Dim countCell As New TableCell
            Dim countLabel As New Label

            countCell.VerticalAlign = VerticalAlign.Middle
            countCell.HorizontalAlign = HorizontalAlign.Left
            countCell.Wrap = False
            countCell.Style.Add("height", "22px")
            countCell.CssClass = "DtlLabel"
            countCell.Style.Add("width", "80%")
            countLabel.ID = "countLabel"
            countLabel.Font.Bold = True
            countLabel.ForeColor = Drawing.Color.DarkSlateBlue
            If Session("gLang") = "E" Then
                countLabel.Text = "Total Record :"
            ElseIf Session("gLang") = "C" Then
                countLabel.Text = "記錄總數 :"
            End If
            countCell.Controls.Add(countLabel)
            printRow.Cells.Add(countCell)

            If printTbl_ADDED = False Then
                printTbl.Rows.Add(printRow)
                dtdiv.Controls.Add(printTbl)
                printTbl_ADDED = True
                printTbl.Visible = False
            End If


            If srch_table.Rows(0).Item("srch_custom_ctrl_links").ToString.Trim <> "" Then
                Try
                    Dim links As String() = srch_table.Rows(0).Item("srch_custom_ctrl_links").ToString.Trim.Split(",")

                    Dim printCtrlCell As New TableCell
                    printCtrlCell.VerticalAlign = VerticalAlign.Middle
                    printCtrlCell.HorizontalAlign = HorizontalAlign.Left
                    printCtrlCell.Wrap = False
                    printCtrlCell.Style.Add("height", "25px")
                    printCtrlCell.CssClass = "DtlLabel"

                    Dim ctrlTable As New Table
                    Dim ctrlRow As New TableRow

                    ctrlTable.CellPadding = 0
                    ctrlTable.CellSpacing = 0

                    ctrlTable.Rows.Add(ctrlRow)

                    For i As Integer = 0 To links.Count - 1
                        Dim ctrlPlaceHolder As New PlaceHolder

                        Dim ctrlCell As New TableCell
                        ctrlCell.CssClass = "DtlLabel"
                        ctrlCell.Wrap = False
                        ctrlCell.Style.Add("text-align", "justify")

                        Dim userCtrl As Control = Page.LoadControl(links(i))
                        'userCtrl.GetType().GetProperty("Message").SetValue(userCtrl, "Hello", Nothing)


                        ctrlPlaceHolder.Controls.Add(userCtrl)

                        ctrlCell.Controls.Add(ctrlPlaceHolder)
                        ctrlRow.Cells.Add(ctrlCell)
                    Next

                    printCtrlCell.Controls.Add(ctrlTable)
                    printRow.Cells.Add(printCtrlCell)

                    If printTbl_ADDED = False Then
                        printTbl.Rows.Add(printRow)
                        dtdiv.Controls.Add(printTbl)
                        printTbl_ADDED = True
                        printTbl.Visible = False
                    End If

                Catch ex As Exception

                End Try

            End If

            If srch_table.Rows(0).Item("srch_rpt_show_download_yn").ToString.Trim = "Y" Then
                If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                    Dim saveLbl As New Label

                    saveLbl.Font.Bold = True

                    If Session("gLang") = "E" Then
                        saveLbl.Text = "Download As : "
                    ElseIf Session("gLang") = "C" Then
                        saveLbl.Text = "下載為 : "
                    End If

                    Dim saveasDDL As New DropDownList
                    Dim SAVEASSQL As String = "SELECT COLC_CODE, COLC_ENG_VALUE AS COLC_VALUE FROM WMS_COL_CODE WHERE COLC_TABCOL = 'REPORT_FILE_TYPE' "

                    If srch_table.Rows(0).Item("srch_rpt_filetype_restriction").ToString.Trim = "1" Then
                        SAVEASSQL = SAVEASSQL & " AND COLC_CODE = 'EXCEL'"
                    ElseIf srch_table.Rows(0).Item("srch_rpt_filetype_restriction").ToString.Trim = "2" Then
                        SAVEASSQL = SAVEASSQL & " AND COLC_CODE = 'PDF'"
                    End If

                    SAVEASSQL = SAVEASSQL & " ORDER BY COLC_DISPLAY_SEQ "

                    saveasDDL.Font.Size = 9
                    saveasDDL.ID = "ddlSaveAs"

                    'uiFun.load_dropdownBy_ColCode(saveasDDL, "CMS_REPORT.FILE_TYPE", Session("gLang"), , , srch_table.Rows(0).Item("srch_rpt_default_filetype").ToString.Trim)
                    uiFun.load_dropdown(saveasDDL, SAVEASSQL, "COLC_CODE", "COLC_VALUE", "", Session("gSelectLabel"), srch_table.Rows(0).Item("srch_rpt_default_filetype").ToString.Trim, True)

                    printCell.Controls.Add(saveLbl)
                    printCell.Controls.Add(saveasDDL)
                    printCell.Controls.Add(New LiteralControl("&nbsp;"))

                    printRow.Cells.Add(printCell)

                    If printTbl_ADDED = False Then
                        printTbl.Rows.Add(printRow)
                        dtdiv.Controls.Add(printTbl)
                        printTbl_ADDED = True
                        printTbl.Visible = False
                    End If
                End If

                If menu_code = "INQ_001" Then
                    Dim ReportExpSTKbtn As New Button
                    ReportExpSTKbtn.ID = "btnExpReport"
                    ReportExpSTKbtn.CssClass = "all_button"

                    If Session("gLang") = "E" Then
                        ReportExpSTKbtn.Text = "Export to Excel"
                    ElseIf Session("gLang") = "C" Then
                        ReportExpSTKbtn.Text = "導出到Excel"
                    End If

                    Dim sm1 As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                    sm1.RegisterPostBackControl(ReportExpSTKbtn)

                    AddHandler ReportExpSTKbtn.Click, AddressOf click_Exp_STKReport

                    printCell.Controls.Add(ReportExpSTKbtn)
                    printRow.Cells.Add(printCell)
                End If

                If menu_code = "OB_DO" Then

                    Dim Reportbtn As New Button
                    Reportbtn.ID = "btnReport"
                    Reportbtn.CssClass = "all_button"

                    If Session("gLang") = "E" Then
                        Reportbtn.Text = "Pick List Report"
                    ElseIf Session("gLang") = "C" Then
                        Reportbtn.Text = "下載"
                    End If

                    Dim sm1 As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                    sm1.RegisterPostBackControl(Reportbtn)

                    AddHandler Reportbtn.Click, AddressOf click_report

                    printCell.Controls.Add(Reportbtn)
                    printRow.Cells.Add(printCell)
                End If

                If menu_code = "OB_DO" Then
                    Dim ReportSSbtn As New Button
                    ReportSSbtn.ID = "btnSSReport"
                    ReportSSbtn.CssClass = "all_button"

                    If Session("gLang") = "E" Then
                        ReportSSbtn.Text = "Short Ship Report"
                    ElseIf Session("gLang") = "C" Then
                        ReportSSbtn.Text = "Short Ship Report"
                    End If

                    Dim sSm1 As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                    sSm1.RegisterPostBackControl(ReportSSbtn)

                    AddHandler ReportSSbtn.Click, AddressOf click_SSreport

                    printCell.Controls.Add(ReportSSbtn)
                    printRow.Cells.Add(printCell)
                End If

                Dim downloadbtn As New Button

                downloadbtn.ID = "btnDownload"
                downloadbtn.CssClass = "all_button"

                If Session("gLang") = "E" Then
                    downloadbtn.Text = "Download"
                ElseIf Session("gLang") = "C" Then
                    downloadbtn.Text = "下載"
                End If

                Dim sm As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                sm.RegisterPostBackControl(downloadbtn)

                AddHandler downloadbtn.Click, AddressOf click_download

                printCell.Controls.Add(downloadbtn)
                printRow.Cells.Add(printCell)

                If printTbl_ADDED = False Then
                    printTbl.Rows.Add(printRow)
                    dtdiv.Controls.Add(printTbl)
                    printTbl_ADDED = True
                    printTbl.Visible = False
                Else
                    printCell.Controls.Add(New LiteralControl("&nbsp;"))
                End If
            End If


            If srch_table.Rows(0).Item("srch_rpt_show_print_yn").ToString.Trim = "Y" Then
                Dim printbtn As New Button

                printbtn.ID = "btnPrint"
                printbtn.CssClass = "all_button"

                If Session("gLang") = "E" Then
                    printbtn.Text = "Print"
                ElseIf Session("gLang") = "C" Then
                    printbtn.Text = "列印"
                End If

                Dim sm As System.Web.UI.ScriptManager = System.Web.UI.ScriptManager.GetCurrent(Page)
                sm.RegisterPostBackControl(printbtn)

                AddHandler printbtn.Click, AddressOf click_print

                printCell.Controls.Add(printbtn)
                printRow.Cells.Add(printCell)

                If printTbl_ADDED = False Then
                    printTbl.Rows.Add(printRow)
                    dtdiv.Controls.Add(printTbl)
                    printTbl_ADDED = True
                    printTbl.Visible = False
                End If
            End If

        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Private Sub GenerateGV()

        Try
            Dim field As String

            If Session("gLang") = "C" Then
                field = "srcl_chi_label"
            Else
                field = "srcl_label"
            End If

            Dim dtGrid As New GridView
            Dim dtSQL As String = "select *, 'Y' as fldo_field_option from wms_search_col where fun_code = '" & gU.dbEncode(menu_code) & "' order by srcl_list_seq, cold_tabcol"

            If ViewState("ctemp") IsNot Nothing AndAlso ViewState("ctemp") Then
                dtSQL = "select wms_search_col.*, ISNULL(wms_field_option.fldo_field_option, 'Y') as fldo_field_option " &
                        "from wms_search_col left outer join wms_field_option on " &
                        "wms_search_col.cold_tabcol = wms_field_option.FLDO_FIELD_NAME " &
                        "and wms_search_col.FUN_CODE = wms_field_option.FUN_CODE " &
                        "and wms_field_option.imp_code= '" & gU.dbEncode(Session("imp_code")) & "' " &
                        "and wms_field_option.Storer_code= '" & gU.dbEncode(ViewState("sc")) & "' " &
                        "where wms_search_col.fun_code = '" & gU.dbEncode(menu_code) & "' " &
                        "order by srcl_list_seq, cold_tabcol"
            End If

            Dim dtTbl As New DataTable
            Dim PKList As String() = Nothing
            Dim tempPK As String = ""
            Dim mc_linkList As New ArrayList

            dtTbl = gDB.getDataTable(dtSQL)

            ViewState("g_list_col_table") = dtTbl

            dtGrid.ID = menu_code & "_GridView"

            If srch_table.Rows(0).Item("SRCH_TYPE").ToString = "LOOKUP" Then
                If srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" Then
                    'Dim templateField As New TemplateField

                    'templateField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                    'templateField.HeaderTemplate = New GridViewLabelTemplate(ListItemType.Header, "chkSelectAll", "", "MCHECKBOX")
                    'templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnCB", "Select", "MCHECKBOX")

                    'templateField.ItemStyle.CssClass = "GV"

                    'dtGrid.Columns.Add(templateField)
                Else
                    Dim templateField As New TemplateField

                    templateField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                    templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnSelect", "btnSelect", "HTMLINPUTBUTTON", , , , , , "Select")
                    templateField.ItemStyle.CssClass = "GV"

                    dtGrid.Columns.Add(templateField)
                End If
            End If

            Dim urlKey As String = ""
            Dim startIndex As Integer = 0
            Dim keyCodeArray As New ArrayList
            Dim keyArray As New ArrayList

            For i As Integer = 0 To dtTbl.Rows.Count - 1
                Dim nFieldName As String = ""
                Dim ntempField As String()

                If dtTbl.Rows(i).Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso dtTbl.Rows(i).Item("srcl_app_custom_ctrl_links").ToString.Trim = "" Then
                    If dtTbl.Rows(i).Item("cold_tabcol").ToString.Trim.Contains(".") Then
                        ntempField = dtTbl.Rows(i).Item("cold_tabcol").ToString.Trim.Split(".")
                        nFieldName = ntempField(1).ToString.Trim
                    Else
                        nFieldName = dtTbl.Rows(i).Item("cold_tabcol").ToString.Trim
                    End If

                    If dtTbl.Rows(i).Item("SRCL_KEY_YN").ToString.Trim.ToUpper = "Y" Then
                        keyArray.Add(nFieldName)
                        If tempPK <> "" Then tempPK = tempPK & ","

                        tempPK = tempPK & nFieldName

                        If gU.decodeNullOrEmpty(dtTbl.Rows(i).Item("srcl_key_code").ToString.Trim.ToLower, "") <> "" Then
                            If urlKey <> "" Then urlKey = urlKey & "&"
                            urlKey = urlKey & dtTbl.Rows(i).Item("srcl_key_code").ToString.Trim.ToLower & "={" & startIndex.ToString & "}"
                            keyCodeArray.Add(dtTbl.Rows(i).Item("srcl_key_code").ToString.Trim.ToLower)
                            startIndex += 1
                        End If
                    End If
                End If
            Next

            ViewState("keyCodeArray") = keyCodeArray
            ViewState("keyArray") = keyArray

            If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
                srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then

                Dim templateField As New TemplateField

                templateField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                templateField.HeaderTemplate = New GridViewLabelTemplate(ListItemType.Header, "chkSelectAll", "", "MCHECKBOX")

                Dim cbSelected As New HiddenField
                Dim cbUnSelected As New HiddenField
                Dim toggle_selectAll As New HiddenField

                cbSelected.ID = "cbSelected"
                srchForm.Controls.Add(cbSelected)

                cbUnSelected.ID = "cbUnSelected"
                srchForm.Controls.Add(cbUnSelected)

                toggle_selectAll.ID = "toggle_selectAll"
                srchForm.Controls.Add(toggle_selectAll)

                Dim cbJs As String = "chkSelected(this,'"
                Dim tmpStr As String = ""

                For idx As Integer = 0 To keyArray.Count - 1
                    If tmpStr <> "" Then tmpStr += "#"
                    tmpStr += "{" & idx & "}"
                Next

                cbJs += tmpStr & "');"

                If tmpStr = "" Then
                    templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnCB", "Select", "MCHECKBOX")
                Else
                    templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnCB", "Select", "MCHECKBOX", , , , "onclick", cbJs)
                End If

                templateField.ItemStyle.CssClass = "GV"
                templateField.HeaderStyle.Width = 25
                templateField.ItemStyle.Width = 25

                dtGrid.Columns.Add(templateField)
            End If

            Dim customFieldList As New ArrayList

            For x As Integer = 0 To dtTbl.Rows.Count - 1
                Dim TableName As String = ""
                Dim FieldName As String = ""
                Dim tempField As String()

                If dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Contains(".") Then
                    tempField = dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Split(".")
                    TableName = tempField(0).ToString.Trim
                    FieldName = tempField(1).ToString.Trim
                Else
                    FieldName = dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim
                End If

                Dim jsKey As String = ""
                Dim jsVal As String = ""

                If dtTbl.Rows(x).Item("srcl_js_attributes_key").ToString.Trim.ToLower <> "" Then
                    jsKey = dtTbl.Rows(x).Item("srcl_js_attributes_key").ToString.Trim.ToLower
                End If

                If dtTbl.Rows(x).Item("srcl_js_attributes_value").ToString.Trim <> "" Then
                    jsVal = dtTbl.Rows(x).Item("srcl_js_attributes_value").ToString.Trim
                End If

                If jsKey = "" Then jsVal = ""

                Dim ltemplateField As New TemplateField

                If dtTbl.Rows(x).Item("srcl_list_col_width").ToString <> "" Then
                    ltemplateField.HeaderStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                    ltemplateField.ItemStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                End If

                ltemplateField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                ltemplateField.HeaderText = dtTbl.Rows(x).Item(field).ToString.Trim
                ltemplateField.ItemStyle.CssClass = "GV"

                If dtTbl.Rows(x).Item("srcl_list_seq").ToString <> "0" AndAlso dtTbl.Rows(x).Item("srcl_list_seq").ToString <> "" AndAlso dtTbl.Rows(x).Item("fldo_field_option").ToString = "Y" Then
                    If dtTbl.Rows(x).Item("srcl_custom_field_yn").ToString.Trim.ToUpper = "Y" Then
                        If dtTbl.Rows(x).Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                            Dim UP As New UpdatePanel
                            UP.ID = "UP_" & dtGrid.ID

                            If dtdiv.FindControl(UP.ID) Is Nothing Then

                                AddHandler UP.Unload, AddressOf UpdatePanel_Unload
                                dtdiv.Controls.Add(UP)
                            End If

                            ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "CUSTOMCONTROL", , , ,
                                                                                                        jsKey, jsVal, dtTbl.Rows(x).Item("srcl_custom_field_text").ToString, dtTbl.Rows(x).Item("srcl_app_custom_ctrl_links").ToString.Trim, UP)

                        ElseIf dtTbl.Rows(x).Item("srcl_href_yn").ToString.Trim.ToUpper = "Y" Then
                            Dim href_Link As String = ""

                            If jsKey.ToLower <> "onclick" Then
                                If urlKey <> "" AndAlso dtTbl.Rows(x).Item("srcl_href_request_para_yn").ToString.Trim.ToUpper = "Y" Then
                                    If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Contains("?") Then
                                        If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Substring(dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Length - 1, 1) = "&" Then
                                            href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & urlKey
                                        Else
                                            href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "&" & urlKey
                                        End If
                                    Else
                                        href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "?" & urlKey
                                    End If
                                ElseIf dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim <> "" Then
                                    href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim
                                Else
                                    href_Link = "#"
                                End If
                            Else
                                href_Link = "#"
                            End If

                            ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "HTMLANCHOR", , , href_Link, jsKey, jsVal, dtTbl.Rows(x).Item("srcl_custom_field_text").ToString)

                        ElseIf dtTbl.Rows(x).Item("srcl_edit_style").ToString.Trim.ToUpper = "B" Then
                            ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "HTMLINPUTBUTTON", , , , jsKey, jsVal, dtTbl.Rows(x).Item("srcl_custom_field_text").ToString)
                        Else
                            ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "LABEL", , , , , dtTbl.Rows(x).Item("srcl_custom_field_text").ToString)
                        End If

                        ltemplateField.SortExpression = FieldName
                        dtGrid.Columns.Add(ltemplateField)

                        customFieldList.Add(dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"))
                    ElseIf dtTbl.Rows(x).Item("srcl_edit_style").ToString.Trim.ToUpper = "B" Then
                        ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "HTMLINPUTBUTTON", , , , jsKey, jsVal, dtTbl.Rows(x).Item("srcl_button_text").ToString)
                    Else
                        Dim colCode As New DataTable
                        Dim colSQL As String = "select 1 from wms_col_code where colc_tabcol = '" & gU.dbEncode(dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim) & "'"

                        colCode = gDB.getDataTable(colSQL)

                        If colCode.Rows.Count = 0 Then
                            If dtTbl.Rows(x).Item("srcl_href_yn").ToString.Trim.ToUpper = "Y" Then
                                Dim netField As String() = tempPK.Split(",")

                                Dim Hyper As New HyperLinkField

                                If dtTbl.Rows(x).Item("srcl_list_col_width").ToString <> "" Then
                                    Hyper.HeaderStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                                    Hyper.ItemStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                                End If

                                Hyper.HeaderStyle.HorizontalAlign = HorizontalAlign.Left

                                Hyper.DataTextField = FieldName

                                If urlKey <> "" AndAlso dtTbl.Rows(x).Item("srcl_href_request_para_yn").ToString.Trim.ToUpper = "Y" Then
                                    Hyper.DataNavigateUrlFields = netField

                                    If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Contains("?") Then
                                        If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Substring(dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Length - 1, 1) = "&" Then
                                            Hyper.DataNavigateUrlFormatString = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & urlKey
                                        Else
                                            Hyper.DataNavigateUrlFormatString = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "&" & urlKey
                                        End If

                                    Else
                                        Hyper.DataNavigateUrlFormatString = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "?" & urlKey
                                    End If
                                ElseIf dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim <> "" Then
                                    Hyper.DataNavigateUrlFormatString = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim
                                Else
                                    Hyper.DataNavigateUrlFormatString = "#"
                                End If

                                Hyper.HeaderText = dtTbl.Rows(x).Item(field).ToString.Trim

                                Hyper.ItemStyle.CssClass = "GV"

                                Hyper.SortExpression = FieldName

                                dtGrid.Columns.Add(Hyper)
                            Else
                                'Dim bField As New BoundField

                                'If dtTbl.Rows(x).Item("srcl_list_col_width").ToString <> "" Then
                                '    bField.HeaderStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                                '    bField.ItemStyle.Width = dtTbl.Rows(x).Item("srcl_list_col_width").ToString
                                'End If

                                'bField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left

                                'bField.DataField = FieldName

                                'bField.HeaderText = dtTbl.Rows(x).Item(field).ToString.Trim

                                'bField.ItemStyle.CssClass = "GV"

                                'bField.SortExpression = FieldName

                                'dtGrid.Columns.Add(bField)
                                'SRCL_EDIT_STYLE

                                ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "LABEL", , , , , dtTbl.Rows(x).Item("srcl_custom_field_text").ToString)

                                dtGrid.Columns.Add(ltemplateField)
                            End If
                        Else
                            If dtTbl.Rows(x).Item("srcl_href_yn").ToString.Trim.ToUpper = "Y" Then
                                Dim href_Link As String = ""

                                If urlKey <> "" AndAlso dtTbl.Rows(x).Item("srcl_href_request_para_yn").ToString.Trim.ToUpper = "Y" Then
                                    If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Contains("?") Then
                                        If dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Substring(dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim.Length - 1, 1) = "&" Then
                                            href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & urlKey
                                        Else
                                            href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "&" & urlKey
                                        End If
                                    Else
                                        href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim & "?" & urlKey
                                    End If
                                ElseIf dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim <> "" Then
                                    href_Link = dtTbl.Rows(x).Item("srcl_href_url").ToString.Trim
                                Else
                                    href_Link = "#"
                                End If

                                ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "HTMLANCHOR", , , href_Link, jsKey, jsVal)
                            Else
                                ltemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, dtTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Replace(".", "|"), "", "LABEL")
                            End If

                            ltemplateField.SortExpression = FieldName
                            dtGrid.Columns.Add(ltemplateField)
                        End If
                    End If

                End If
            Next

            ViewState("customFieldList") = customFieldList

            If tempPK <> "" Then
                PKList = tempPK.Split(",")
                dtGrid.DataKeyNames = PKList
            End If

            If srch_table.Rows(0).Item("srch_show_delete_button_yn").ToString.ToUpper.Trim = "Y" Then
                Dim templateField As New TemplateField

                templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnDelete", , "DELETEBUTTON", , , , , , "Delete")
                templateField.ItemStyle.CssClass = "GV"

                dtGrid.Columns.Add(templateField)
            End If

            dtGrid.Style.Add("width", "100%")

            If srch_table.Rows(0).Item("srch_disable_paging_yn").ToString.ToUpper.Trim = "Y" Then
                dtGrid.AllowPaging = False
            Else
                dtGrid.AllowPaging = True

                Dim pagerDDL As New DropDownList
                pagerDDL.ID = "pager_select"
                pagerDDL.AutoPostBack = True
                pagerDDL.Width = 45
                pagerDDL.Style("vertical-align") = "middle"

                AddHandler pagerDDL.SelectedIndexChanged, AddressOf gvrsList_PagerIndexChanged

                Dim firstBtn As New ImageButton
                Dim previousBtn As New ImageButton
                Dim nextBtn As New ImageButton
                Dim lastBtn As New ImageButton

                firstBtn.ID = "pager_first"
                firstBtn.ImageUrl = "images/arrow_frist.png"
                firstBtn.Height = 24
                firstBtn.Width = 24
                firstBtn.Style("vertical-align") = "middle"

                AddHandler firstBtn.Click, AddressOf gvrsList_PagerIndexFirst

                previousBtn.ID = "pager_previous"
                previousBtn.ImageUrl = "images/arrow_previous.png"
                previousBtn.Height = 24
                previousBtn.Width = 24
                previousBtn.Style("vertical-align") = "middle"

                AddHandler previousBtn.Click, AddressOf gvrsList_PagerIndexBackward

                nextBtn.ID = "pager_next"
                nextBtn.ImageUrl = "images/arrow_next.png"
                nextBtn.Height = 24
                nextBtn.Width = 24
                nextBtn.Style("vertical-align") = "middle"

                AddHandler nextBtn.Click, AddressOf gvrsList_PagerIndexForward

                lastBtn.ID = "pager_last"
                lastBtn.ImageUrl = "images/arrow_last.png"
                lastBtn.Height = 24
                lastBtn.Width = 24
                lastBtn.Style("vertical-align") = "middle"

                AddHandler lastBtn.Click, AddressOf gvrsList_PagerIndexLast

                dtGrid.PagerTemplate = New GridViewPagerTemplate(ListItemType.Pager, pagerDDL, previousBtn, nextBtn, firstBtn, lastBtn)
            End If

            dtGrid.Font.Name = "Arial"
            dtGrid.Font.Overline = False
            dtGrid.AutoGenerateColumns = False

            dtGrid.PageSize = gU.decodeEmptyCInt(srch_table.Rows(0).Item("srch_page_size").ToString, "20")

            dtGrid.BorderStyle = BorderStyle.None
            dtGrid.BorderWidth = 1
            dtGrid.CellPadding = 3
            dtGrid.CaptionAlign = TableCaptionAlign.Top

            dtGrid.RowStyle.BorderColor = Drawing.Color.Silver
            dtGrid.RowStyle.BorderStyle = BorderStyle.Solid
            dtGrid.RowStyle.BorderWidth = 1

            dtGrid.HeaderStyle.BorderColor = Drawing.Color.White
            dtGrid.HeaderStyle.BorderStyle = BorderStyle.Inset
            dtGrid.HeaderStyle.BorderWidth = 1
            dtGrid.HeaderStyle.Font.Name = "Arial"
            dtGrid.HeaderStyle.Font.Bold = True
            dtGrid.HeaderStyle.Font.Size = 9
            dtGrid.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
            dtGrid.HeaderStyle.VerticalAlign = VerticalAlign.Middle
            dtGrid.HeaderStyle.Wrap = True
            dtGrid.HeaderStyle.CssClass = "DtlLabel"

            dtGrid.FooterStyle.BackColor = Drawing.Color.White
            dtGrid.SelectedRowStyle.Font.Bold = True
            dtGrid.SelectedRowStyle.ForeColor = Drawing.Color.White

            dtGrid.PagerStyle.Font.Size = 8
            dtGrid.PagerStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#4C60B6")
            dtGrid.PagerStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#006699")
            dtGrid.PagerStyle.HorizontalAlign = HorizontalAlign.Left
            dtGrid.PagerStyle.BorderColor = Drawing.Color.Transparent
            dtGrid.PagerStyle.BorderWidth = 1
            dtGrid.PagerStyle.Font.Bold = True
            dtGrid.PagerStyle.Font.Name = "Arial"
            dtGrid.PagerStyle.Font.Strikeout = False

            'If srch_table.Rows(0).Item("srch_allow_col_sort_yn").ToString.ToUpper.Trim <> "Y" Then
            '    dtGrid.AllowSorting = True
            'Else
            '    dtGrid.AllowSorting = False
            'End If

            AddHandler dtGrid.RowDeleting, AddressOf dtGrid_OnRowDeleting
            AddHandler dtGrid.RowUpdating, AddressOf dtGrid_OnRowUpdating
            AddHandler dtGrid.RowDataBound, AddressOf gvrsList_RowDataBound
            AddHandler dtGrid.DataBound, AddressOf gvrsList_DataBound

            AddHandler dtGrid.PageIndexChanging, AddressOf gvrsList_PageIndexChanging
            AddHandler dtGrid.Sorting, AddressOf gvrsList_Sorting

            If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim <> "Y" Then
                If Session("gLang") = "C" Then
                    dtGrid.EmptyDataText = "找不到記錄."
                Else
                    dtGrid.EmptyDataText = "No Record Found."
                End If

                dtdiv.Controls.Add(dtGrid)
            Else
                Dim trdtHD As New TableRow
                Dim tddtHD As New TableCell
                Dim BarDtlbl As New Label
                BarDtlbl.ID = "ItemBar_Title"

                If Session("gLang") = "E" Then
                    BarDtlbl.Text = "Please Select The Items"
                ElseIf Session("gLang") = "C" Then
                    BarDtlbl.Text = "請選擇項目"
                End If

                BarDtlbl.Font.Bold = True
                tddtHD.CssClass = "TITLE"

                tddtHD.Controls.Add(BarDtlbl)

                trdtHD.Controls.Add(tddtHD)
                Dim dtTable As New Table
                dtTable.CellPadding = 1
                dtTable.CellSpacing = 1

                dtTable.Style.Add("width", "100%")
                dtTable.HorizontalAlign = HorizontalAlign.Left

                dtTable.ID = "detailTbl"

                dtTable.Controls.Add(trdtHD)

                Dim gvRow As New TableRow
                Dim gvCell As New TableCell

                gvCell.Controls.Add(dtGrid)
                gvRow.Controls.Add(gvCell)

                dtTable.Controls.Add(gvRow)

                dtTable.Visible = False

                dtdiv.Controls.Add(dtTable)
            End If
        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Private Sub GenerateMasterGV()
        Try

            If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                Dim field As String

                If Session("gLang") = "C" Then
                    field = "srcl_chi_label"
                Else
                    field = "srcl_label"
                End If

                Dim MasterGrid As New GridView

                MasterGrid.ID = menu_code & "_MasterGridView"

                Dim MasterTbl As New DataTable
                Dim PKList As String() = Nothing
                Dim tempPK As String = ""
                Dim mc_linkList As New ArrayList

                Dim mSQL As String = "select * from wms_search_col where fun_code = '" & menu_code & "' order by srcl_master_list_seq"

                MasterTbl = gDB.getDataTable(mSQL)

                ViewState("g_master_list_col_table") = MasterTbl

                Dim urlKey As String = ""
                Dim startIndex As Integer = 0

                For i As Integer = 0 To MasterTbl.Rows.Count - 1
                    Dim nFieldName As String = ""
                    Dim ntempField As String()

                    If MasterTbl.Rows(i).Item("cold_tabcol").ToString.Trim.Contains(".") Then
                        ntempField = MasterTbl.Rows(i).Item("cold_tabcol").ToString.Trim.Split(".")
                        nFieldName = ntempField(1).ToString.Trim
                    Else
                        nFieldName = MasterTbl.Rows(i).Item("cold_tabcol").ToString.Trim
                    End If

                    If MasterTbl.Rows(i).Item("SRCL_KEY_YN").ToString.Trim.ToUpper = "Y" Then
                        If tempPK <> "" Then tempPK = tempPK & ","

                        tempPK = tempPK & nFieldName

                        If gU.decodeNullOrEmpty(MasterTbl.Rows(i).Item("srcl_key_code").ToString.Trim.ToLower, "") <> "" Then
                            If urlKey <> "" Then urlKey = urlKey & "&"
                            urlKey = urlKey & MasterTbl.Rows(i).Item("srcl_key_code").ToString.Trim.ToLower & "={" & startIndex.ToString & "}"
                            startIndex += 1
                        End If
                    End If


                Next

                For x As Integer = 0 To MasterTbl.Rows.Count - 1
                    Dim TableName As String = ""
                    Dim FieldName As String = ""
                    Dim tempField As String()

                    If MasterTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Contains(".") Then
                        tempField = MasterTbl.Rows(x).Item("cold_tabcol").ToString.Trim.Split(".")
                        TableName = tempField(0).ToString.Trim
                        FieldName = tempField(1).ToString.Trim
                    Else
                        FieldName = MasterTbl.Rows(x).Item("cold_tabcol").ToString.Trim
                    End If

                    If MasterTbl.Rows(x).Item("srcl_master_list_seq").ToString <> "0" And MasterTbl.Rows(x).Item("srcl_master_list_seq").ToString <> "" Then
                        If MasterTbl.Rows(x).Item("srcl_mc_href_yn").ToString.Trim.ToUpper = "Y" Then

                            Dim ntemplateField As New TemplateField

                            If MasterTbl.Rows(x).Item("srcl_list_col_width").ToString <> "" Then
                                ntemplateField.HeaderStyle.Width = MasterTbl.Rows(x).Item("srcl_list_col_width").ToString
                                ntemplateField.ItemStyle.Width = MasterTbl.Rows(x).Item("srcl_list_col_width").ToString
                            End If

                            ntemplateField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                            ntemplateField.HeaderTemplate = New GridViewLabelTemplate(ListItemType.Header, , , "LINKBUTTON", "", MasterTbl.Rows(x).Item(field).ToString.Trim)
                            ntemplateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, FieldName, , "LINKBUTTON", FieldName)

                            MasterGrid.Columns.Add(ntemplateField)

                            mc_LList.Add(FieldName)

                        Else
                            Dim bField As New BoundField

                            If MasterTbl.Rows(x).Item("srcl_list_col_width").ToString <> "" Then
                                bField.HeaderStyle.Width = MasterTbl.Rows(x).Item("srcl_list_col_width").ToString
                                bField.ItemStyle.Width = MasterTbl.Rows(x).Item("srcl_list_col_width").ToString
                            End If

                            bField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left

                            bField.DataField = FieldName

                            bField.HeaderText = MasterTbl.Rows(x).Item(field).ToString.Trim

                            bField.ItemStyle.CssClass = "GV"

                            MasterGrid.Columns.Add(bField)
                        End If
                    End If
                Next

                If tempPK <> "" Then
                    PKList = tempPK.Split(",")
                    MasterGrid.DataKeyNames = PKList
                End If

                If srch_table.Rows(0).Item("srch_show_delete_button_yn").ToString.ToUpper.Trim = "Y" Then
                    Dim templateField As New TemplateField

                    templateField.ItemTemplate = New GridViewLabelTemplate(ListItemType.Item, "btnDelete", , "DELETEBUTTON", , , , , , "Delete")

                    MasterGrid.Columns.Add(templateField)
                End If

                'ViewState("g_mc_linkList") = mc_linkList
                'mc_linkList.Clear()

                If srch_table.Rows(0).Item("srch_disable_paging_yn").ToString.ToUpper.Trim = "Y" Then
                    MasterGrid.AllowPaging = False
                Else
                    MasterGrid.AllowPaging = True

                    If Session("SEARCH_SESSION_PAGE_MasterGV_PAGEINDEX") IsNot Nothing Then
                        MasterGrid.PageIndex = Session("SEARCH_SESSION_PAGE_MasterGV_PAGEINDEX")
                    End If

                End If

                MasterGrid.Font.Name = "Arial"
                MasterGrid.Font.Overline = False
                MasterGrid.AutoGenerateColumns = False
                MasterGrid.PageSize = 20

                MasterGrid.BorderStyle = BorderStyle.None
                MasterGrid.BorderWidth = 1
                MasterGrid.CellPadding = 3
                MasterGrid.CaptionAlign = TableCaptionAlign.Top

                MasterGrid.RowStyle.BorderColor = Drawing.Color.Silver
                MasterGrid.RowStyle.BorderStyle = BorderStyle.NotSet
                MasterGrid.RowStyle.BorderWidth = 1

                MasterGrid.HeaderStyle.BorderColor = Drawing.Color.White
                MasterGrid.HeaderStyle.BorderStyle = BorderStyle.Inset
                MasterGrid.HeaderStyle.BorderWidth = 1
                MasterGrid.HeaderStyle.Font.Name = "Arial"
                MasterGrid.HeaderStyle.Font.Bold = True
                MasterGrid.HeaderStyle.Font.Size = 9
                MasterGrid.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
                MasterGrid.HeaderStyle.VerticalAlign = VerticalAlign.Bottom
                MasterGrid.HeaderStyle.Wrap = True
                MasterGrid.HeaderStyle.CssClass = "DtlLabel"

                MasterGrid.FooterStyle.BackColor = Drawing.Color.White
                MasterGrid.SelectedRowStyle.Font.Bold = True
                MasterGrid.SelectedRowStyle.ForeColor = Drawing.Color.White

                MasterGrid.PagerStyle.Font.Size = 8
                MasterGrid.PagerStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#4C60B6")
                MasterGrid.PagerStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#006699")
                MasterGrid.PagerStyle.HorizontalAlign = HorizontalAlign.Left
                MasterGrid.PagerStyle.BorderColor = Drawing.Color.Transparent
                MasterGrid.PagerStyle.BorderWidth = 1
                MasterGrid.PagerStyle.Font.Bold = True
                MasterGrid.PagerStyle.Font.Name = "Arial"
                MasterGrid.PagerStyle.Font.Strikeout = False
                If Session("gLang") = "C" Then
                    MasterGrid.EmptyDataText = "找不到記錄."
                Else
                    MasterGrid.EmptyDataText = "No Record Found."
                End If


                MasterGrid.Style.Add("width", "96%")

                'AddHandler MasterGrid.RowDeleting, AddressOf dtGrid_OnRowDeleting
                AddHandler MasterGrid.RowUpdating, AddressOf MasterGrid_OnRowUpdating
                AddHandler MasterGrid.RowDataBound, AddressOf masterGv_RowDataBound
                AddHandler MasterGrid.PageIndexChanging, AddressOf MasterList_PageIndexChanging

                '<table border="0" cellspacing="0" cellpadding="0">
                '        <tr>
                '            <td width="100%" class="TITLE">
                '                <b><asp:Label ID="lbl_ImageHd" runat="server" /></b>
                '            </td>
                '        </tr>
                '    </table>

                Dim trBarHD As New TableRow
                Dim tdBarHD As New TableCell
                Dim BarHDlbl As New Label
                BarHDlbl.ID = "HeaderBar_Title"

                If Session("gLang") = "E" Then
                    BarHDlbl.Text = "Please Select The Code"
                ElseIf Session("gLang") = "C" Then
                    BarHDlbl.Text = "請選擇編號"
                End If

                BarHDlbl.Font.Bold = True
                tdBarHD.CssClass = "TITLE"

                tdBarHD.Controls.Add(BarHDlbl)
                trBarHD.Controls.Add(tdBarHD)

                Dim mTbl As New Table
                mTbl.CellPadding = 1
                mTbl.CellSpacing = 1

                mTbl.Style.Add("width", "98%")
                mTbl.HorizontalAlign = HorizontalAlign.Left

                mTbl.ID = "masterTbl"

                mTbl.Controls.Add(trBarHD)

                Dim divRow As New TableRow
                Dim divCell As New TableCell

                divCell.BorderColor = System.Drawing.ColorTranslator.FromHtml("#BBBBBB")
                divCell.BorderWidth = 1
                divCell.BorderStyle = BorderStyle.Solid

                Dim headerDiv As New HtmlGenericControl("div")
                headerDiv.ID = "mGVHDDiv"
                headerDiv.Style.Add("overflow", "auto")
                headerDiv.Style.Add("width", "100%")
                headerDiv.Style.Add("border", "0px Inset")

                Dim dataDiv As New HtmlGenericControl("div")
                dataDiv.ID = "mGVDiv"
                dataDiv.Style.Add("overflow", "auto")
                dataDiv.Style.Add("border", "0px Inset")
                dataDiv.Style.Add("border-color", "#BBBBBB")
                dataDiv.Style.Add("width", "100%")

                dataDiv.Style.Add("height", "240px")
                dataDiv.Attributes.Add("onscroll", "Onscrollfnction();")

                dataDiv.Controls.Add(MasterGrid)
                divCell.Controls.Add(headerDiv)
                divCell.Controls.Add(dataDiv)
                divRow.Controls.Add(divCell)

                mTbl.Controls.Add(divRow)

                'mPanel.Controls.Add(MasterGrid)
                'mGVDiv.Controls.Add(MasterGrid)
                mdiv.Controls.Add(mTbl)
            End If
        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Private Sub GenerateObjectSQL(ByRef SCString As String, ByVal tblRow As DataRow, ByVal srcl_edit_style As String, ByVal oper As String,
                                    ByVal ctrlName As String, ByVal FieldName As String, ByVal symbol As String,
                                    Optional ByVal org_edit_style As String = "",
                                    Optional ByRef form_post As FormPosting = Nothing,
                                    Optional ByRef ctlArrayList As Dictionary(Of String, String()) = Nothing,
                                    Optional ByRef isAbort As Boolean = False,
                                    Optional ByVal upperEntry As Boolean = False)

        Dim item_srch_sql As String = ""
        Dim isSelectField As Boolean = False

        If srch_table.Rows(0).Item("SRCH_SELECT_FIELD_YN").ToString.Trim.ToUpper = "Y" Then
            isSelectField = True
        End If

        If gU.decodeNullOrEmpty(tblRow.Item("srcl_hide_entry_yn").ToString, "") = "Y" Then
            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                    ctlArrayList.Add(FieldName, New String() {ctrlName, FieldName, "", ""})
                End If
            End If

            Exit Sub
        End If

        Select Case srcl_edit_style.ToUpper.Trim
            Case "T", "U"
                If gU.decodeNullOrEmpty(tblRow.Item("srcl_range_yn").ToString.Trim.ToUpper, "N") = "N" Then
                    Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                    If validateObj(tmpTextbox.Text.Trim, srcl_edit_style, tblRow.Item("srcl_datatype").ToString) Then
                        Session("SEARCH_SESSION_PAGE_" & tmpTextbox.ID) = tmpTextbox.Text.Trim

                        If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                        End If

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : " & gU.decodeNullOrEmpty(tmpTextbox.Text.Trim, "All")

                        If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                            If tmpTextbox.Text.Trim = "" Then
                                If Session("gLang") = "E" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                                ElseIf Session("gLang") = "C" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                                End If

                                isAbort = True
                                Exit Sub
                            End If
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" Then
                            Select Case oper.ToUpper
                                Case "P"
                                    If upperEntry Then
                                        item_srch_sql = "UPPER(" & FieldName & ") LIKE N'" & gU.dbEncode(tmpTextbox.Text.Trim.ToUpper) & "%' "
                                    Else
                                        item_srch_sql = FieldName & " LIKE N'" & gU.dbEncode(tmpTextbox.Text.Trim) & "%' "
                                    End If
                                Case "="
                                    item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpTextbox.Text.Trim) & "' "
                                Case "LIKE"
                                    item_srch_sql = FieldName & " LIKE "

                                    If upperEntry Then
                                        item_srch_sql = "UPPER(" & FieldName & ") LIKE N'%" & gU.dbEncode(tmpTextbox.Text.Trim.ToUpper) & "%' "
                                    Else
                                        item_srch_sql += "N'%" & gU.dbEncode(tmpTextbox.Text.Trim) & "%' "
                                    End If
                                Case Else
                                    item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpTextbox.Text.Trim) & "' "
                            End Select

                            If tmpTextbox.Text.Trim <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                                SCString = APPSQL(SCString, item_srch_sql, symbol)
                            End If
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpTextbox.ID, tmpTextbox.Text.Trim)
                        End If

                        If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                            If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                                Dim pass_obj_sql As String = ""

                                If tmpTextbox.Text.Trim <> "" Then
                                    pass_obj_sql = item_srch_sql
                                End If

                                ctlArrayList.Add(FieldName, New String() {tmpTextbox.ID, FieldName, tmpTextbox.Text.Trim, pass_obj_sql})
                            End If
                        End If
                    Else
                        isAbort = True
                        Exit Sub
                    End If
                Else
                    Dim tmpFromTextbox As TextBox = DirectCast(hddiv.FindControl("FR_" & ctrlName), TextBox)
                    Dim tmpToTextbox As TextBox = DirectCast(hddiv.FindControl("TO_" & ctrlName), TextBox)

                    If validateObj(tmpFromTextbox.Text.Trim, srcl_edit_style, tblRow.Item("srcl_datatype").ToString) AndAlso validateObj(tmpToTextbox.Text.Trim, srcl_edit_style, tblRow.Item("srcl_datatype").ToString) Then
                        Session("SEARCH_SESSION_PAGE_" & tmpFromTextbox.ID) = tmpFromTextbox.Text.Trim
                        Session("SEARCH_SESSION_PAGE_" & tmpToTextbox.ID) = tmpToTextbox.Text.Trim

                        If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                        End If

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " :"

                        If tmpFromTextbox.Text.Trim = "" AndAlso tmpToTextbox.Text.Trim = "" Then
                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " All"
                        Else
                            If tmpFromTextbox.Text.Trim <> "" Then
                                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " From " & tmpFromTextbox.Text.Trim
                            End If

                            If tmpToTextbox.Text.Trim <> "" Then
                                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " To " & tmpToTextbox.Text.Trim
                            End If
                        End If

                        If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                            If tmpFromTextbox.Text.Trim = "" AndAlso tmpToTextbox.Text.Trim = "" Then
                                If Session("gLang") = "E" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                                ElseIf Session("gLang") = "C" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                                End If

                                isAbort = True
                                Exit Sub
                            End If
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                            tmpFromTextbox.Text.Trim <> "" AndAlso
                            tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                            item_srch_sql += FieldName & " >= '" & tmpFromTextbox.Text.Trim & "' "
                            'item_srch_sql = APPSQL(item_srch_sql, FieldName & " >= '" & tmpFromTextbox.Text.Trim & "' ", symbol)
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                            tmpToTextbox.Text.Trim <> "" AndAlso
                            tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                            If item_srch_sql <> "" Then item_srch_sql += "AND "
                            item_srch_sql += FieldName & " <= '" & gU.dbEncode(tmpToTextbox.Text.Trim) & "' "
                        End If

                        If item_srch_sql <> "" Then
                            SCString = APPSQL(SCString, item_srch_sql, symbol)
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpFromTextbox.ID, tmpFromTextbox.Text.Trim)
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpToTextbox.ID, tmpToTextbox.Text.Trim)
                        End If

                        If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                            If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                                Dim pass_obj_sql As String = item_srch_sql

                                ctlArrayList.Add(FieldName, New String() {tmpFromTextbox.ID & "#" & tmpToTextbox.ID, FieldName, tmpFromTextbox.Text.Trim & "#" & tmpToTextbox.Text.Trim, pass_obj_sql})
                            End If
                        End If
                    Else
                        isAbort = True
                        Exit Sub
                    End If
                End If

                'Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                'Session("SEARCH_SESSION_PAGE_" & tmpTextbox.ID) = tmpTextbox.Text.Trim

                'If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                '    If tmpTextbox.Text.Trim = "" Then
                '        If Session("gLang") = "E" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('Please Fill-in Search Criteria!\r\n""" & tblRow.Item("srcl_label").ToString & """ cannot be empty.');</script>")
                '        ElseIf Session("gLang") = "C" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('" & tblRow.Item("srcl_chi_label").ToString & "不能留空');</script>")
                '        End If

                '        isAbort = True
                '        Exit Sub
                '    End If
                'End If

                'Select Case oper.ToUpper
                '    Case "="
                '        item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpTextbox.Text.Trim) & "' "
                '    Case "LIKE"
                '        'item_srch_sql = FieldName & " LIKE '%" & gU.dbEncode(tmpTextbox.Text.Trim) & "%' "
                '        item_srch_sql = "UPPER(" & FieldName & ") LIKE '%" & gU.dbEncode(UCase(tmpTextbox.Text.Trim)) & "%' "
                '    Case Else
                '        Exit Sub
                'End Select

                'If tmpTextbox.Text.Trim <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "Y" Then
                '    SCString = APPSQL(SCString, item_srch_sql, symbol)
                'End If

                'If form_post IsNot Nothing Then
                '    form_post.Add(tmpTextbox.ID, tmpTextbox.Text.Trim)
                'End If

                'If tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper = "Y" Then
                '    If ctlArrayList IsNot Nothing Then
                '        ctlArrayList.Add(New String() {tmpTextbox.ID, FieldName, tmpTextbox.Text.Trim})
                '    End If
                'End If
            Case "D"
                If gU.decodeNullOrEmpty(tblRow.Item("srcl_range_yn").ToString.Trim.ToUpper, "Y") = "N" Then
                    Dim tmpDateTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                    If validateObj(tmpDateTextbox.Text.Trim, srcl_edit_style) Then
                        Session("SEARCH_SESSION_PAGE_" & tmpDateTextbox.ID) = tmpDateTextbox.Text.Trim

                        If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                        End If

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : " & gU.decodeNullOrEmpty(tmpDateTextbox.Text.Trim, "All")

                        If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                            If tmpDateTextbox.Text.Trim = "" Then
                                If Session("gLang") = "E" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                                ElseIf Session("gLang") = "C" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                                End If

                                isAbort = True
                                Exit Sub
                            End If
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                            tmpDateTextbox.Text.Trim <> "" AndAlso
                            tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                            item_srch_sql = FieldName & " >= CONVERT(DATETIME, '" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "'," & gU.getConfig("DDFORMATNO") & ") AND " & FieldName & " < CONVERT(DATETIME,'" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "'," & gU.getConfig("DDFORMATNO") & ") +1 "
                            'item_srch_sql = APPSQL(item_srch_sql, FieldName & " >= to_date('" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "','" & gU.getConfig("DDFORMAT") & "') ", symbol)
                            'item_srch_sql = APPSQL(item_srch_sql, FieldName & " < to_date('" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "','" & gU.getConfig("DDFORMAT") & "') +1 ", symbol)
                        End If

                        If item_srch_sql <> "" Then
                            'If SCString <> "" Then SCString += " "
                            'SCString += item_srch_sql
                            SCString = APPSQL(SCString, item_srch_sql, symbol)
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpDateTextbox.ID, tmpDateTextbox.Text.Trim)
                        End If

                        If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                            If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                                Dim pass_obj_sql As String = item_srch_sql

                                ctlArrayList.Add(FieldName, New String() {tmpDateTextbox.ID, FieldName, tmpDateTextbox.Text.Trim, pass_obj_sql})
                            End If
                        End If
                    Else
                        isAbort = True
                        Exit Sub
                    End If
                Else
                    Dim tmpFromTextbox As TextBox = DirectCast(hddiv.FindControl("FR_" & ctrlName), TextBox)
                    Dim tmpToTextbox As TextBox = DirectCast(hddiv.FindControl("TO_" & ctrlName), TextBox)

                    If validateObj(tmpFromTextbox.Text.Trim, srcl_edit_style) AndAlso validateObj(tmpToTextbox.Text.Trim, srcl_edit_style) Then
                        Session("SEARCH_SESSION_PAGE_" & tmpFromTextbox.ID) = tmpFromTextbox.Text.Trim
                        Session("SEARCH_SESSION_PAGE_" & tmpToTextbox.ID) = tmpToTextbox.Text.Trim

                        If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                        End If

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " :"

                        If tmpFromTextbox.Text.Trim = "" AndAlso tmpFromTextbox.Text.Trim = "" Then
                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " All"
                        Else
                            If tmpFromTextbox.Text.Trim <> "" Then
                                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " From " & tmpFromTextbox.Text.Trim
                            End If

                            If tmpFromTextbox.Text.Trim <> "" AndAlso tmpToTextbox.Text.Trim <> "" Then
                                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " To " & tmpToTextbox.Text.Trim
                            ElseIf tmpToTextbox.Text.Trim <> "" Then
                                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += " Till " & tmpToTextbox.Text.Trim
                            End If
                        End If

                        If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                            If tmpFromTextbox.Text.Trim = "" AndAlso tmpToTextbox.Text.Trim = "" Then
                                If Session("gLang") = "E" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Select Date!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                                ElseIf Session("gLang") = "C" Then
                                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                                End If

                                isAbort = True
                                Exit Sub
                            End If
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                            tmpFromTextbox.Text.Trim <> "" AndAlso
                            tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                            item_srch_sql += FieldName & " >= CONVERT(DATETIME,'" & gU.dbEncode(tmpFromTextbox.Text.Trim) & "'," & gU.getConfig("DDFORMATNO") & ") "
                            'item_srch_sql = APPSQL(item_srch_sql, FieldName & " >= to_date('" & gU.dbEncode(tmpFromTextbox.Text.Trim) & "','" & gU.getConfig("DDFORMAT") & "') ", symbol)
                        End If

                        If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                            tmpToTextbox.Text.Trim <> "" AndAlso
                            tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                            If item_srch_sql <> "" Then item_srch_sql += "AND "
                            item_srch_sql += FieldName & " < CONVERT(DATETIME,'" & gU.dbEncode(tmpToTextbox.Text.Trim) & "'," & gU.getConfig("DDFORMATNO") & ") +1 "
                            'item_srch_sql = APPSQL(item_srch_sql, FieldName & " < to_date('" & gU.dbEncode(tmpToTextbox.Text.Trim) & "','" & gU.getConfig("DDFORMAT") & "') +1 ", symbol)
                        End If

                        If item_srch_sql <> "" Then
                            'If SCString <> "" Then SCString += " "
                            'SCString += item_srch_sql
                            SCString = APPSQL(SCString, item_srch_sql, symbol)
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpFromTextbox.ID, tmpFromTextbox.Text.Trim)
                        End If

                        If form_post IsNot Nothing Then
                            form_post.Add(tmpToTextbox.ID, tmpToTextbox.Text.Trim)
                        End If

                        If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                            If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                                Dim pass_obj_sql As String = item_srch_sql

                                ctlArrayList.Add(FieldName, New String() {tmpFromTextbox.ID & "#" & tmpToTextbox.ID, FieldName, tmpFromTextbox.Text.Trim & "#" & tmpToTextbox.Text.Trim, pass_obj_sql})
                            End If
                        End If
                    Else
                        isAbort = True
                        Exit Sub
                    End If

                End If
                'If tblRow.Item("srcl_date_range_yn").ToString.Trim.ToUpper = "N" Then
                '    Dim tmpDateTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                '    If validateObj(tmpDateTextbox.Text.Trim, srcl_edit_style) Then
                '        Session("SEARCH_SESSION_PAGE_" & tmpDateTextbox.ID) = tmpDateTextbox.Text.Trim

                '        If tmpDateTextbox.Text.Trim <> "" Then
                '            If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim <> "" Then
                '                If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("SESSION") Then
                '                    Dim sessionName As String = ""

                '                    sessionName = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Substring( _
                '                                    tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"), _
                '                                    tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Length - tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"))

                '                    sessionName = sessionName.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", "")

                '                    Dim addStr As String = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Replace("SESSION(""" & sessionName & """)", "")

                '                    SCString = APPSQL(SCString, addStr & " " & Session(sessionName), symbol)

                '                ElseIf tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("REQUEST") Then

                '                Else
                '                    SCString = APPSQL(SCString, tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim, symbol)

                '                    SCString = String.Format(SCString, tmpDateTextbox.Text.Trim)
                '                End If
                '            End If
                '        Else
                '            If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                '                If tmpDateTextbox.Text.Trim = "" Then
                '                    If Session("gLang") = "E" Then
                '                        ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                                       "<script>alert('Please Select Date!\r\n""" & tblRow.Item("srcl_label").ToString & """ cannot be empty.');</script>")
                '                    ElseIf Session("gLang") = "C" Then
                '                        ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                                       "<script>alert('" & tblRow.Item("srcl_chi_label").ToString & "不能留空');</script>")
                '                    End If

                '                    isAbort = True
                '                    Exit Sub
                '                End If
                '            End If
                '        End If

                '        If tmpDateTextbox.Text.Trim <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "Y" AndAlso tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" Then
                '            SCString = APPSQL(SCString, FieldName & " >= to_date('" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "','" & Cache("DDFORMAT") & "') ", symbol)
                '            SCString = APPSQL(SCString, FieldName & " < to_date('" & gU.dbEncode(tmpDateTextbox.Text.Trim) & "','" & Cache("DDFORMAT") & "') +1 ", symbol)
                '        End If

                '        If form_post IsNot Nothing Then
                '            form_post.Add(tmpDateTextbox.ID, tmpDateTextbox.Text.Trim)
                '        End If

                '        If tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper = "Y" Then
                '            If ctlArrayList IsNot Nothing Then
                '                ctlArrayList.Add(New String() {tmpDateTextbox.ID, FieldName, tmpDateTextbox.Text.Trim})
                '            End If
                '        End If
                '    Else
                '        isAbort = True
                '        Exit Sub
                '    End If
                'Else
                '    Dim tmpFromTextbox As TextBox = DirectCast(hddiv.FindControl("FR_" & ctrlName), TextBox)
                '    Dim tmpToTextbox As TextBox = DirectCast(hddiv.FindControl("TO_" & ctrlName), TextBox)

                '    If validateObj(tmpFromTextbox.Text.Trim, srcl_edit_style) AndAlso validateObj(tmpToTextbox.Text.Trim, srcl_edit_style) Then
                '        Session("SEARCH_SESSION_PAGE_" & tmpFromTextbox.ID) = tmpFromTextbox.Text.Trim
                '        Session("SEARCH_SESSION_PAGE_" & tmpToTextbox.ID) = tmpToTextbox.Text.Trim

                '        If tmpFromTextbox.Text.Trim <> "" OrElse tmpToTextbox.Text.Trim <> "" Then
                '            If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim <> "" Then
                '                If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("SESSION") Then
                '                    Dim sessionName As String = ""

                '                    sessionName = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Substring( _
                '                                    tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"), _
                '                                    tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Length - tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"))

                '                    sessionName = sessionName.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", "")

                '                    Dim addStr As String = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Replace("SESSION(""" & sessionName & """)", "")

                '                    SCString = APPSQL(SCString, addStr & " " & Session(sessionName), symbol)

                '                ElseIf tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("REQUEST") Then

                '                Else
                '                    SCString = APPSQL(SCString, tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim, symbol)

                '                End If
                '            End If
                '        ElseIf tmpFromTextbox.Text.Trim = "" AndAlso tmpToTextbox.Text.Trim = "" Then
                '            If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                '                If tmpFromTextbox.Text.Trim = "" And tmpToTextbox.Text.Trim = "" Then
                '                    If Session("gLang") = "E" Then
                '                        ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                                       "<script>alert('Please Select Date!\r\n""" & tblRow.Item("srcl_label").ToString & """ cannot be empty.');</script>")
                '                    ElseIf Session("gLang") = "C" Then
                '                        ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                                       "<script>alert('" & tblRow.Item("srcl_chi_label").ToString & "不能留空');</script>")
                '                    End If

                '                    isAbort = True
                '                    Exit Sub
                '                End If
                '            End If
                '        End If

                '        If tmpFromTextbox.Text.Trim <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "Y" AndAlso tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" Then
                '            SCString = APPSQL(SCString, FieldName & " >= to_date('" & gU.dbEncode(tmpFromTextbox.Text.Trim) & "','" & Cache("DDFORMAT") & "') ", symbol)
                '        End If

                '        If tmpToTextbox.Text.Trim <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "Y" AndAlso tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" Then
                '            SCString = APPSQL(SCString, FieldName & " < to_date('" & gU.dbEncode(tmpToTextbox.Text.Trim) & "','" & Cache("DDFORMAT") & "') +1 ", symbol)
                '        End If

                '        If form_post IsNot Nothing Then
                '            form_post.Add(tmpFromTextbox.ID, tmpFromTextbox.Text.Trim)
                '        End If

                '        If form_post IsNot Nothing Then
                '            form_post.Add(tmpToTextbox.ID, tmpToTextbox.Text.Trim)
                '        End If

                '        If tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper = "Y" Then
                '            If ctlArrayList IsNot Nothing Then
                '                ctlArrayList.Add(New String() {tmpFromTextbox.ID & "#" & tmpToTextbox.ID, FieldName, tmpFromTextbox.Text.Trim & "#" & tmpToTextbox.Text.Trim})
                '            End If
                '        End If
                '    Else
                '        isAbort = True
                '        Exit Sub
                '    End If

                'End If

            Case "C", "R"
                Dim tmpCBValue As String = ""

                Dim tmpCheckBoxList As New CheckBoxList
                tmpCheckBoxList = DirectCast(hddiv.FindControl(ctrlName), CheckBoxList)

                If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                End If

                If tmpCheckBoxList.SelectedIndex <> -1 Then
                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                        item_srch_sql = FieldName & " IN (" & uiFun.getValueFromCB(tmpCheckBoxList) & ") "
                        'item_srch_sql = APPSQL(item_srch_sql, FieldName & " IN (" & uiFun.getValueFromCB(tmpCheckBoxList) & ") ", symbol)
                    End If

                    If item_srch_sql <> "" Then
                        'If SCString <> "" Then SCString += " "
                        'SCString += item_srch_sql
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                    Dim tmpstr As String = ""
                    Dim tS As String = ""
                    Dim addSQL As String = ""

                    For i As Integer = 0 To tmpCheckBoxList.Items.Count - 1
                        If tmpCheckBoxList.Items(i).Selected = True Then
                            If tmpstr <> "" Then tmpstr = tmpstr & ","
                            tmpstr += tmpCheckBoxList.Items(i).Value
                        End If
                    Next

                    Session("SEARCH_SESSION_PAGE_" & tmpCheckBoxList.ID) = tmpstr

                    If tmpstr.Trim <> "" Then
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : " & tmpstr
                    Else
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : All"
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpCheckBoxList.ID, tmpstr)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = item_srch_sql

                            ctlArrayList.Add(FieldName, New String() {tmpCheckBoxList.ID, FieldName, tmpstr, pass_obj_sql})
                        End If
                    End If
                Else
                    Session("SEARCH_SESSION_PAGE_" & tmpCheckBoxList.ID) = ""

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : All"

                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        If Session("gLang") = "E" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Tick Checkbox!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be un-selected.');", True)
                        ElseIf Session("gLang") = "C" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                        End If

                        isAbort = True
                        Exit Sub
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpCheckBoxList.ID, "")
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            ctlArrayList.Add(FieldName, New String() {tmpCheckBoxList.ID, FieldName, "", ""})
                        End If
                    End If
                End If
                'Dim tmpCBValue As String = ""

                'Dim tmpCheckBoxList As New CheckBoxList
                'tmpCheckBoxList = DirectCast(hddiv.FindControl(ctrlName), CheckBoxList)

                'If tmpCheckBoxList.SelectedIndex <> -1 Then
                '    Dim tmpstr As String = ""
                '    Dim tS As String = ""
                '    Dim addSQL As String = ""

                '    If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim <> "" Then
                '        If tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("SESSION") Then
                '            Dim sessionName As String = ""

                '            sessionName = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Substring( _
                '                            tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"), _
                '                            tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Length - tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.IndexOf("SESSION"))

                '            sessionName = sessionName.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", "")

                '            Dim addStr As String = tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Replace("SESSION(""" & sessionName & """)", "")

                '            SCString = APPSQL(SCString, addStr & " " & Session(sessionName), symbol)

                '        ElseIf tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim.Contains("REQUEST") Then

                '        Else
                '            SCString = APPSQL(SCString, tblRow.Item("srcl_add_where_clause").ToString.ToUpper.Trim, symbol)
                '        End If
                '    Else
                '        SCString = APPSQL(SCString, "UPPER(" & FieldName & ") IN (" & uiFun.getValueFromCB(tmpCheckBoxList) & ") ", symbol)
                '    End If

                '    For i As Integer = 0 To tmpCheckBoxList.Items.Count - 1
                '        If tmpCheckBoxList.Items(i).Selected = True Then
                '            If tmpstr <> "" Then tmpstr = tmpstr & ","
                '            tmpstr = tmpstr & tmpCheckBoxList.Items(i).Value

                '            If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                '                tS = "select colc_add_where_sql from wms_col_code where " & _
                '                        "UPPER(colc_tabcol) = '" & tblRow.Item("cold_tabcol").ToString.Trim.ToUpper & "' and " & _
                '                        "UPPER(colc_code) = '" & tmpCheckBoxList.Items(i).Value.ToUpper.ToString & "' "
                '                addSQL = dbFun.getValueFromSQL(tS).ToUpper.ToString

                '                If addSQL <> "" Then
                '                    SCString = SCString & " " & addSQL
                '                End If
                '            End If
                '        End If
                '    Next

                '    Session("SEARCH_SESSION_PAGE_" & tmpCheckBoxList.ID) = tmpstr

                '    If form_post IsNot Nothing Then
                '        form_post.Add(tmpCheckBoxList.ID, tmpstr)
                '    End If
                'Else
                '    Session("SEARCH_SESSION_PAGE_" & tmpCheckBoxList.ID) = ""

                '    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then

                '        If Session("gLang") = "E" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('Please Tick Checkbox!\r\n""" & tblRow.Item("srcl_label").ToString & """ cannot be un-selected.');</script>")
                '        ElseIf Session("gLang") = "C" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('" & tblRow.Item("srcl_chi_label").ToString & "不能留空');</script>")
                '        End If

                '        isAbort = True
                '        Exit Sub
                '    End If

                '    If form_post IsNot Nothing Then
                '        form_post.Add(tmpCheckBoxList.ID, "")
                '    End If
                'End If
            Case "K"
                Dim tmpCheckBox As New CheckBox
                tmpCheckBox = DirectCast(hddiv.FindControl(ctrlName), CheckBox)

                If tmpCheckBox.Checked = True Then
                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then

                        item_srch_sql = FieldName & "  = '1' "
                    End If

                    If item_srch_sql <> "" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                    Session("SEARCH_SESSION_PAGE_" & tmpCheckBox.ID) = tmpCheckBox.Checked

                    If tmpCheckBox.Checked Then
                        If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                            Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                        End If

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : Yes"
                    Else
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : No"
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpCheckBox.ID, tmpCheckBox.Checked)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = item_srch_sql

                            ctlArrayList.Add(FieldName, New String() {tmpCheckBox.ID, FieldName, True, pass_obj_sql})
                        End If
                    End If
                Else
                    Session.Remove("SEARCH_SESSION_PAGE_" & tmpCheckBox.ID)

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                            gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                            " : No"

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            ctlArrayList.Add(FieldName, New String() {tmpCheckBox.ID, FieldName, "", ""})
                        End If
                    End If
                End If

            Case "S"
                Dim tmpDDLValue As String = ""

                Dim tmpDropDownList As New DropDownList
                tmpDropDownList = DirectCast(hddiv.FindControl(ctrlName), DropDownList)

                If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                End If

                If tmpDropDownList.SelectedValue.ToString <> "" Then
                    Session("SEARCH_SESSION_PAGE_" & tmpDropDownList.ID) = tmpDropDownList.SelectedItem.Value

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                        gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                        " : " & gU.decodeNullOrEmpty(tmpDropDownList.SelectedValue.ToString, "All")

                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" Then
                        Select Case oper.ToUpper
                            Case "P"
                                If upperEntry Then
                                    item_srch_sql = "UPPER(" & FieldName & ") LIKE N'" & gU.dbEncode(tmpDropDownList.SelectedValue.ToUpper) & "%' "
                                Else
                                    item_srch_sql = FieldName & " LIKE N'" & gU.dbEncode(tmpDropDownList.SelectedValue) & "%' "
                                End If
                            Case "="
                                item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpDropDownList.SelectedValue) & "' "
                            Case "LIKE"
                                item_srch_sql = FieldName & " LIKE "

                                If upperEntry Then
                                    item_srch_sql = "UPPER(" & FieldName & ") LIKE N'%" & gU.dbEncode(tmpDropDownList.SelectedValue.ToUpper) & "%' "
                                Else
                                    item_srch_sql += "N'%" & gU.dbEncode(tmpDropDownList.SelectedValue) & "%' "
                                End If
                            Case Else
                                item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpDropDownList.SelectedValue) & "' "
                        End Select

                        If tmpDropDownList.SelectedValue <> "" AndAlso tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                            SCString = APPSQL(SCString, item_srch_sql, symbol)
                        End If
                    End If

                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        srcl_edit_style.ToUpper.Trim = "L" Then

                        Dim tS As String = "select colc_sql from lm_col_code where " &
                                        "cold_tabcol = '" & tblRow.Item("cold_tabcol").ToString.Trim & "' and " &
                                        "colc_code = '" & tmpDropDownList.SelectedValue.ToString & "' "

                        Dim addSQL As String = dbFun.getValueFromSQL(tS).ToString

                        If addSQL <> "" Then
                            SCString = SCString.Trim & " AND " & addSQL
                        End If
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpDropDownList.ID, tmpDropDownList.SelectedItem.Value)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = item_srch_sql

                            ctlArrayList.Add(FieldName, New String() {tmpDropDownList.ID, FieldName, tmpDropDownList.SelectedItem.Value, pass_obj_sql})
                        End If
                    End If
                Else
                    Session("SEARCH_SESSION_PAGE_" & tmpDropDownList.ID) = ""

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                        gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                        " : All"

                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        If Session("gLang") = "E" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Select Item!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be un-selected.');", True)
                        ElseIf Session("gLang") = "C" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                        End If

                        isAbort = True
                        Exit Sub
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpDropDownList.ID, "")
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            ctlArrayList.Add(FieldName, New String() {tmpDropDownList.ID, FieldName, "", ""})
                        End If
                    End If
                End If

                'Dim tmpDDLValue As String = ""

                'Dim tmpDropDownList As New DropDownList
                'tmpDropDownList = DirectCast(hddiv.FindControl(ctrlName), DropDownList)

                'If tmpDropDownList.SelectedValue.ToString <> "" Then
                '    Session("SEARCH_SESSION_PAGE_" & tmpDropDownList.ID) = tmpDropDownList.SelectedItem.Value
                '    If tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "Y" Then
                '        SCString = APPSQL(SCString, FieldName & " = '" & gU.dbEncode(tmpDropDownList.SelectedItem.Value) & "' ", symbol)
                '    End If

                '    If form_post IsNot Nothing Then
                '        form_post.Add(tmpDropDownList.ID, tmpDropDownList.SelectedItem.Value)
                '    End If

                '    If tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper = "Y" Then
                '        If ctlArrayList IsNot Nothing Then
                '            ctlArrayList.Add(New String() {tmpDropDownList.ID, FieldName, tmpDropDownList.SelectedItem.Value})
                '        End If
                '    End If
                'Else
                '    Session("SEARCH_SESSION_PAGE_" & tmpDropDownList.ID) = ""

                '    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                '        If Session("gLang") = "E" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('Please Select Item!\r\n""" & tblRow.Item("srcl_label").ToString & """ cannot be un-selected.');</script>")
                '        ElseIf Session("gLang") = "C" Then
                '            ClientScript.RegisterStartupScript(Me.GetType, "ManWarn", _
                '                                           "<script>alert('" & tblRow.Item("srcl_chi_label").ToString & "不能留空');</script>")
                '        End If

                '        isAbort = True
                '        Exit Sub
                '    End If

                '    If form_post IsNot Nothing Then
                '        form_post.Add(tmpDropDownList.ID, "")
                '    End If
                'End If

            Case "L"
                Dim tmpHidden As HiddenField = TryCast(srchForm.FindControl(ctrlName), HiddenField)
                Dim tmpLabelHd As HiddenField = TryCast(srchForm.FindControl("dsphd_" & ctrlName), HiddenField)
                Dim tmpLabel As TextBox = TryCast(hddiv.FindControl("dsp_" & ctrlName), TextBox)

                If tmpHidden IsNot Nothing Then
                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        If tmpHidden.Value.Trim = "" Then
                            If Session("gLang") = "E" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                            ElseIf Session("gLang") = "C" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                            End If

                            isAbort = True
                            Exit Sub
                        End If
                    End If

                    Session("SEARCH_SESSION_PAGE_" & tmpHidden.ID) = tmpHidden.Value

                    If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                    End If

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                        gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                        " : " & gU.decodeNullOrEmpty(tmpHidden.Value, "All")

                    Select Case oper.ToUpper
                        Case "P"
                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'" & gU.dbEncode(tmpHidden.Value.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql = FieldName & " LIKE N'" & gU.dbEncode(tmpHidden.Value.Trim) & "%' "
                            End If
                        Case "="
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpHidden.Value.Trim) & "' "
                        Case "LIKE"
                            item_srch_sql = FieldName & " LIKE "

                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'%" & gU.dbEncode(tmpHidden.Value.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql += "N'%" & gU.dbEncode(tmpHidden.Value.Trim) & "%' "
                            End If
                        Case Else
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpHidden.Value.Trim) & "' "
                    End Select

                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        tmpHidden.Value <> "" AndAlso
                        tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpHidden.ID, tmpHidden.Value)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = ""

                            If tmpHidden.Value <> "" Then
                                pass_obj_sql = item_srch_sql
                            End If

                            ctlArrayList.Add(FieldName, New String() {tmpHidden.ID, FieldName, tmpHidden.Value, pass_obj_sql})
                        End If
                    End If
                Else
                    GenerateObjectSQL(SCString, tblRow, org_edit_style, oper, ctrlName, FieldName, symbol, , form_post, , , upperEntry)
                End If

                If tmpLabelHd IsNot Nothing AndAlso tmpLabel IsNot Nothing Then
                    If tmpLabelHd.Value.Trim <> "" Then
                        ViewState(tmpLabel.ID & "_VIEWSTATE") = tmpLabelHd.Value.Trim
                        Session("SEARCH_SESSION_PAGE_" & tmpLabel.ID) = tmpLabelHd.Value.Trim
                        tmpLabel.Text = tmpLabelHd.Value.Trim
                    ElseIf ViewState(tmpLabel.ID & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(tmpLabel.ID & "_VIEWSTATE").ToString <> "" Then
                        tmpLabelHd.Value = ViewState(tmpLabel.ID & "_VIEWSTATE")
                        Session("SEARCH_SESSION_PAGE_" & tmpLabel.ID) = tmpLabelHd.Value.Trim
                        tmpLabel.Text = tmpLabelHd.Value.Trim
                    End If

                    tmpLabel.ReadOnly = True
                End If

            Case "H"
                Dim tmpHidden As HiddenField = TryCast(hdfieldHolder.FindControl(ctrlName), HiddenField)

                If tmpHidden IsNot Nothing Then
                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        If tmpHidden.Value.Trim = "" Then
                            If Session("gLang") = "E" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                            ElseIf Session("gLang") = "C" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                            End If

                            isAbort = True
                            Exit Sub
                        End If
                    End If

                    Session("SEARCH_SESSION_PAGE_" & tmpHidden.ID) = tmpHidden.Value

                    If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                    End If

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                        gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                        " : " & gU.decodeNullOrEmpty(tmpHidden.Value, "All")

                    Select Case oper.ToUpper
                        Case "P"
                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'" & gU.dbEncode(tmpHidden.Value.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql = FieldName & " LIKE N'" & gU.dbEncode(tmpHidden.Value.Trim) & "%' "
                            End If
                        Case "="
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpHidden.Value.Trim) & "' "
                        Case "LIKE"
                            item_srch_sql = FieldName & " LIKE "

                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'%" & gU.dbEncode(tmpHidden.Value.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql += "N'%" & gU.dbEncode(tmpHidden.Value.Trim) & "%' "
                            End If
                        Case Else
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpHidden.Value.Trim) & "' "
                    End Select

                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        tmpHidden.Value <> "" AndAlso
                        tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpHidden.ID, tmpHidden.Value)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = ""

                            If tmpHidden.Value <> "" Then
                                pass_obj_sql = item_srch_sql
                            End If

                            ctlArrayList.Add(FieldName, New String() {tmpHidden.ID, FieldName, tmpHidden.Value, pass_obj_sql})
                        End If
                    End If
                Else
                    GenerateObjectSQL(SCString, tblRow, org_edit_style, oper, ctrlName, FieldName, symbol, , form_post, , , upperEntry)
                End If

            Case "CSR"
                Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                Session("SEARCH_SESSION_PAGE_" & tmpTextbox.ID) = tmpTextbox.Text.Trim

                If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                End If

                Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                    gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                    " : " & gU.decodeNullOrEmpty(tmpTextbox.Text.Trim, "All")

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    If tmpTextbox.Text.Trim = "" Then
                        If Session("gLang") = "E" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                        ElseIf Session("gLang") = "C" Then
                            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                        End If

                        isAbort = True
                        Exit Sub
                    End If
                End If

                If tmpTextbox.Text.Trim <> "" Then
                    item_srch_sql = gU.splitRngOfString(tmpTextbox.Text.Trim, FieldName, upperEntry)
                End If

                If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                    tmpTextbox.Text.Trim <> "" AndAlso
                    tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                    SCString = APPSQL(SCString, item_srch_sql, symbol)
                End If

                If form_post IsNot Nothing Then
                    form_post.Add(tmpTextbox.ID, tmpTextbox.Text.Trim)
                End If

                If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                    If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                        Dim pass_obj_sql As String = item_srch_sql

                        ctlArrayList.Add(FieldName, New String() {tmpTextbox.ID, FieldName, tmpTextbox.Text, pass_obj_sql})
                    End If
                End If

            Case Else
                Dim tmpTextbox As TextBox = DirectCast(hddiv.FindControl(ctrlName), TextBox)

                If validateObj(tmpTextbox.Text.Trim, srcl_edit_style, tblRow.Item("srcl_datatype").ToString) Then
                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        If tmpTextbox.Text.Trim = "" Then
                            If Session("gLang") = "E" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('Please Fill-in Search Criteria!\r\n""" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & """ cannot be empty.');", True)
                            ElseIf Session("gLang") = "C" Then
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "ManWarn", "alert('" & gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") & "不能留空');", True)
                            End If

                            isAbort = True
                            Exit Sub
                        End If
                    End If

                    Session("SEARCH_SESSION_PAGE_" & tmpTextbox.ID) = tmpTextbox.Text.Trim

                    If Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") IsNot Nothing AndAlso
                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") <> "" Then

                        Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") += vbNewLine

                    End If

                    Session("SEARCH_SESSION_PAGE_CRITERIA_LIST") +=
                        gU.decodeNullOrEmpty(tblRow.Item("srcl_label").ToString.Trim, "") &
                        " : " & gU.decodeNullOrEmpty(tmpTextbox.Text.Trim, "All")

                    Select Case oper.ToUpper
                        Case "P"
                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'" & gU.dbEncode(tmpTextbox.Text.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql = FieldName & " LIKE N'" & gU.dbEncode(tmpTextbox.Text.Trim) & "%' "
                            End If

                        Case "="
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpTextbox.Text.Trim) & "' "
                        Case "LIKE"
                            item_srch_sql = FieldName & " LIKE "

                            If upperEntry Then
                                item_srch_sql = "UPPER(" & FieldName & ") LIKE N'%" & gU.dbEncode(tmpTextbox.Text.Trim.ToUpper) & "%' "
                            Else
                                item_srch_sql += "N'%" & gU.dbEncode(tmpTextbox.Text.Trim) & "%' "
                            End If
                        Case Else
                            item_srch_sql = FieldName & " = '" & gU.dbEncode(tmpTextbox.Text.Trim) & "' "
                    End Select

                    If tblRow.Item("srcl_custom_field_yn").ToString.Trim.ToUpper <> "Y" AndAlso
                        tmpTextbox.Text.Trim <> "" AndAlso
                        tblRow.Item("srcl_generic_pass_value_yn").ToString.Trim.ToUpper <> "N" Then
                        SCString = APPSQL(SCString, item_srch_sql, symbol)
                    End If

                    If form_post IsNot Nothing Then
                        form_post.Add(tmpTextbox.ID, tmpTextbox.Text.Trim)
                    End If

                    If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" OrElse isSelectField Then
                        If ctlArrayList IsNot Nothing AndAlso Not ctlArrayList.ContainsKey(FieldName) Then
                            Dim pass_obj_sql As String = item_srch_sql

                            ctlArrayList.Add(FieldName, New String() {tmpTextbox.ID, FieldName, tmpTextbox.Text, pass_obj_sql})
                        End If
                    End If
                Else
                    isAbort = True
                    Exit Sub
                End If
        End Select
    End Sub

    Protected Sub BindGV(Optional ByVal isMaster As Boolean = False, Optional ByVal masterWhereString As String = "", Optional ByVal clearPageIndex As Boolean = False)
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim srchSQL As String = ""
        Dim whereSQL As String = ""
        Dim criteriaSQL As String = ""
        Dim addSQL As String = ""
        Dim orderbySQL As String = ""
        Dim FieldList As String = ""
        Dim ctrlArray As New Dictionary(Of String, String())
        Dim abortFlag As Boolean = False
        Try
            If masterWhereString <> "" Then
                If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                    generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                    GenericAppMod.ctrlValue = ctrlArray
                    SQLString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
                Else
                    SQLString = generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
                End If
            Else
                If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                    generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                    GenericAppMod.ctrlValue = ctrlArray
                    SQLString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
                Else
                    SQLString = generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
                End If
            End If

            Session("GENERIC_SESSION_SRCH_COMPLETE_SQL") = SQLString
            Session("GENERIC_SESSION_SRCH_SQL") = srchSQL
            Session("GENERIC_SESSION_SRCH_WHERE_SQL") = whereSQL
            Session("GENERIC_SESSION_SRCH_CRITERIA_SQL") = criteriaSQL
            Session("GENERIC_SESSION_SRCH_ADD_SQL") = addSQL
            Session("GENERIC_SESSION_SRCH_ORDERBY_SQL") = orderbySQL
            Session("GENERIC_SESSION_SRCH_SELECT_FIELDS") = FieldList

            If abortFlag Then
                Exit Sub
            Else
                If SQLString <> "" Then
                    dt = gDB.getDataTable(SQLString)
                    'dt = gDB.getDataTable(SQLString.ToUpper)

                    Dim gvrsList As New GridView

                    If isMaster Then
                        gvrsList = DirectCast(mdiv.FindControl(menu_code & "_MasterGridView"), GridView)
                        mdiv.Visible = True

                        Dim GVL As New GridView
                        GVL = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
                        GVL.DataSource = Nothing
                        GVL.DataBind()
                        DirectCast(dtdiv.FindControl("detailTbl"), Table).Visible = False
                    Else
                        gvrsList = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
                        If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                            DirectCast(dtdiv.FindControl("detailTbl"), Table).Visible = True
                            gvrsList.EmptyDataText = "No Items Found."
                        End If

                        'If srch_table.Rows(0).Item("srch_rpt_show_print_yn").ToString.Trim = "Y" OrElse _
                        '    srch_table.Rows(0).Item("srch_rpt_show_download_yn").ToString.Trim = "Y" OrElse _
                        '    srch_table.Rows(0).Item("srch_custom_ctrl_links").ToString.Trim <> "" Then
                        '    If TryCast(dtdiv.FindControl("printTable"), Table) IsNot Nothing Then
                        '        DirectCast(dtdiv.FindControl("printTable"), Table).Visible = True
                        '        ViewState("isShowPrintTable") = True
                        '    End If
                        'End If

                        If TryCast(dtdiv.FindControl("printTable"), Table) IsNot Nothing Then
                            DirectCast(dtdiv.FindControl("printTable"), Table).Visible = True
                            ViewState("isShowPrintTable") = True
                        End If

                        If srch_table.Rows(0).Item("srch_custom_ctrl_links").ToString.Trim <> "" Then
                            If TryCast(dtdiv.FindControl("ctrlTable"), Table) IsNot Nothing Then
                                DirectCast(dtdiv.FindControl("ctrlTable"), Table).Visible = True
                                ViewState("isShowCtrlTable") = True
                            End If
                        End If
                    End If

                    If clearPageIndex Then
                        gvrsList.PageIndex = 0
                    Else
                        If remain_status Then
                            If Not Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") Is Nothing Then
                                gvrsList.PageIndex = Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX")
                            End If
                        End If
                    End If

                    If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
                        srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
                        Session("SEARCH_SESSION_PAGE_GV_DT") = dt
                    End If
                    'ViewState("dt") = dt
                    gvrsList.DataSource = dt
                    gvrsList.DataBind()
                    'countLabel

                    If TryCast(dtdiv.FindControl("printTable"), Table) IsNot Nothing Then
                        If TryCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label) IsNot Nothing Then
                            If dt.Rows.Count = 1 Then
                                If Session("gLang") = "C" Then
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text = "找到共 " & dt.Rows.Count & " 記錄 -"
                                Else
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text = "Total " & dt.Rows.Count & " record found -"
                                End If

                            Else
                                If Session("gLang") = "C" Then
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text = "找到共 " & dt.Rows.Count & " 記錄 -"
                                Else
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text = "Total " & dt.Rows.Count & " records found"
                                End If

                            End If


                            If gvrsList.PageCount > 1 Then
                                If Session("gLang") = "C" Then
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text += " - " & gvrsList.PageCount & " 頁."
                                Else
                                    DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text += " - " & gvrsList.PageCount & " pages."
                                End If

                            Else
                                DirectCast(DirectCast(dtdiv.FindControl("printTable"), Table).FindControl("countLabel"), Label).Text += "."
                            End If
                        End If
                    End If

                    If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
                        srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
                        Session("SEARCH_SESSION_PAGE_GV_DT") = dt
                        Call RePopulateCheckBoxes()
                    End If

                    If isMaster Then
                        If gvrsList.Rows.Count > 0 Then
                            ClientScript.RegisterStartupScript(Me.GetType(), "CreateGridHeader", "<script>CreateGridHeader('mGVDiv', '" & gvrsList.ClientID & "', 'mGVHDDiv');</script>")
                        End If
                    End If
                    SQLString = ""
                Else
                    If Not NoAccess Then
                        System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: SQL is missing.');", True)
                    End If

                    Exit Sub
                End If
            End If
        Catch ex As Exception
            If Session("usr_id") = "OTSADMIN" Then
                'Dim msglog As New PrgmLog("C:\HKE_PROD\HKE_PROD_LOG", "20170321debug.txt")
                '
                'msglog.writeLog("test bindgv error")
                'msglog.writeLog(ex.Message)

                'msglog = Nothing
            End If
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & "\n\r\n\r" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Protected Sub MasterGrid_OnRowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs)

        Dim mGV As GridView = DirectCast(mdiv.FindControl(menu_code & "_MasterGridView"), GridView)

        Dim keyTable As String = srch_table.Rows(0).Item("srch_mtable").ToString
        Dim keySet As New ArrayList
        Dim keyValues As OrderedDictionary = mGV.DataKeys(e.RowIndex).Values
        Dim keySetList As String = ""

        If keyTable <> "" Then
            For x As Integer = 0 To master_list_table.Rows.Count - 1
                If master_list_table.Rows(x).Item("SRCL_KEY_YN").ToString.Trim.ToUpper = "Y" Then
                    keySet.Add(master_list_table.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper)
                End If
            Next

            If keySet.Count = keyValues.Count Then
                For i As Integer = 0 To keySet.Count - 1
                    If keySetList <> "" Then keySetList = keySetList & " AND "
                    keySetList = keySetList & keySet.Item(i) & " = '" & keyValues(i) & "' "
                Next
            End If
        End If
        Call BindGV(False, keySetList)
    End Sub

    Protected Sub dtGrid_OnRowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs)

    End Sub

    Protected Sub dtGrid_OnRowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim keyTable As String = srch_table.Rows(0).Item("srch_mtable").ToString
        Dim keyValue As OrderedDictionary = gvrsList.DataKeys(e.RowIndex).Values
        Dim keyname As New ArrayList
        Dim keynameList As String = ""

        If keyTable <> "" Then
            For x As Integer = 0 To list_table.Rows.Count - 1
                If list_table.Rows(x).Item("SRCL_KEY_YN").ToString.Trim.ToUpper = "Y" Then
                    keyname.Add(list_table.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper)
                End If
            Next

            If keyname.Count = keyValue.Count Then
                For i As Integer = 0 To keyname.Count - 1
                    If keynameList <> "" Then keynameList = keynameList & " AND "
                    keynameList = keynameList & keyname.Item(i) & " = '" & keyValue(i) & "' "
                Next
            End If

            Dim delete_sql As String = ""

            If keynameList <> "" Then

                delete_sql = "DELETE FROM " & srch_table.Rows(0).Item("srch_mtable").ToString & " WHERE " & keynameList

                If srch_table.Rows(0).Item("srch_sessioncol").ToString <> "" Then
                    Dim sessionList As String()
                    Dim tempSessionString As String = ""

                    If srch_table.Rows(0).Item("srch_sessioncol").ToString.Contains(",") Then
                        sessionList = srch_table.Rows(0).Item("srch_sessioncol").ToString.Split(",")

                        For j As Integer = 0 To sessionList.Length - 1
                            tempSessionString = tempSessionString & " AND " & sessionList(j).ToString.ToUpper.Trim & "='" & Session(sessionList(j).ToString.Trim) & "'"
                        Next
                    Else
                        tempSessionString = " AND " & srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim & "='" & Session(srch_table.Rows(0).Item("srch_sessioncol").ToString.Trim) & "'"
                    End If

                    delete_sql = delete_sql & tempSessionString
                End If

            End If

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction
            Dim paP As GlobalDBFunc.DBCmdPara
            Try
                Select Case menu_code
                    Case "MAST_WM"
                        Dim countSQL As String = ""
                        Dim tempCount As Integer = 0
                        Dim lwh_code As String = ""

                        For i As Integer = 0 To keyname.Count - 1
                            If UCase(keyname.Item(i)) = "WMS_WAREHOUSE.WH_CODE" Then
                                lwh_code = keyValue(i)
                                Exit For
                            End If

                        Next
                        If lwh_code <> "" Then
                            countSQL = "Select Count(*) from wms_item_loc_bal where imp_code='" & Session("imp_code") & "' AND ILOC_WH='" & gU.dbEncode(lwh_code) & "'"
                            tempCount = gU.decodeEmptyCInt(dbFun.getValueFromSQL(countSQL, gConn, transaction), 0)

                            If tempCount = 0 Then
                                paP = New GlobalDBFunc.DBCmdPara
                                delete_sql = "Delete from wms_warehouse where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(lwh_code)

                                gDB.amendData(delete_sql, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                delete_sql = "Delete from wms_wh_fl where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(lwh_code)

                                gDB.amendData(delete_sql, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                delete_sql = "Delete from wms_wh_area where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(lwh_code)

                                gDB.amendData(delete_sql, gConn, transaction, paP)

                                paP = New GlobalDBFunc.DBCmdPara
                                delete_sql = "Delete from wms_wh_rack where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(lwh_code)

                                gDB.amendData(delete_sql, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                delete_sql = "Delete from wms_wh_bin where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(lwh_code)

                                gDB.amendData(delete_sql, gConn, transaction, paP)
                            Else
                                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "alertmsg", "alert('The selected Warehouse cannot be deleted!\nLocation Balance exists in this location.');", True)
                                Exit Sub

                            End If
                        End If

                    Case Else
                        If delete_sql <> "" Then gDB.amendData(delete_sql, gConn, transaction)
                End Select


                transaction.Commit()

                Call BindGV()

            Catch ex As Exception
                transaction.Rollback()
                System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(delete_sql.ToUpper) & "\n\r\n\r" & gU.jsString(ex.Message) & "');", True)
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

    Protected Sub masterGv_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                    For i As Integer = 0 To mc_LList.Count - 1
                        Dim sLink As LinkButton
                        sLink = CType(e.Row.FindControl(mc_LList(i)), LinkButton)
                        sLink.Text = DataBinder.Eval(e.Row.DataItem, mc_LList(i)).ToString.Trim
                    Next
                End If
        End Select
    End Sub

    Protected Sub gvrsList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim oGridView As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)

        If oGridView.AllowPaging = True AndAlso oGridView.PageCount > 0 Then
            Dim ddlPager As DropDownList = TryCast(oGridView.BottomPagerRow.FindControl("pager_select"), DropDownList)

            If ddlPager IsNot Nothing Then
                ddlPager.Items.Clear()

                For i As Integer = 1 To oGridView.PageCount
                    ddlPager.Items.Add(New ListItem(i.ToString(), i.ToString()))
                Next

                ddlPager.SelectedValue = oGridView.PageIndex + 1
            End If
        End If

        If Not oGridView.DataSource Is Nothing And
            (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" Or
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") And
            oGridView.Rows.Count <> 0 Then

            Dim chkAll As CheckBox = DirectCast(oGridView.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            chkAll.Attributes("onclick") = "ChgAllChkState(this.checked);"

            Dim ArrayValues As New List(Of String)
            ArrayValues.Add(String.Concat("'", chkAll.ClientID, "'"))

            For Each gvr As GridViewRow In oGridView.Rows
                Dim cb As CheckBox = DirectCast(gvr.FindControl("btnCB"), CheckBox)

                cb.Attributes("onclick") = cb.Attributes("onclick") & "ChgHDState(this);"
                ArrayValues.Add(String.Concat("'", cb.ClientID, "'"))
            Next

            CBCollection.Text = "<script type=""text/javascript"">" & vbCrLf &
                            "<!--" & vbCrLf &
                            String.Concat("var chkArray = new Array(", String.Join(",", ArrayValues.ToArray()), ");") & vbCrLf &
                            "// -->" & vbCrLf &
                            "</script>"

            Dim cbJs As String = ""

            cbJs = String.Concat("var chkArray = new Array(", String.Join(",", ArrayValues.ToArray()), ");")

            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "cbChkArray", cbJs, True)
            'RowsCount.Value = ViewState("dt").Rows.Count

        End If
    End Sub

    Protected Sub gvrsList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim customFieldList As ArrayList = ViewState("customFieldList")

                For Each cell As DataControlFieldCell In e.Row.Cells
                    Dim field As DataControlField = cell.ContainingField

                    If TypeOf field Is HyperLinkField Then
                        Dim td As TableCell = cell

                        If td.Controls.Count > 0 AndAlso TypeOf td.Controls(0) Is HyperLink Then
                            Dim hyperLink As HyperLink = DirectCast(td.Controls(0), HyperLink)
                            Dim hyperLinkField As HyperLinkField = DirectCast(field, HyperLinkField)

                            If Not customFieldList.Contains(hyperLink.ID) Then
                                If Not String.IsNullOrEmpty(hyperLinkField.DataNavigateUrlFormatString) Then
                                    Dim dataUrlFields As String() = New String(hyperLinkField.DataNavigateUrlFields.Length - 1) {}
                                    For j As Integer = 0 To dataUrlFields.Length - 1
                                        Dim obj As Object = DataBinder.Eval(e.Row.DataItem, hyperLinkField.DataNavigateUrlFields(j))
                                        dataUrlFields(j) = HttpUtility.UrlEncode((If(obj Is Nothing, "", obj.ToString())))
                                    Next
                                    hyperLink.NavigateUrl = String.Format(hyperLinkField.DataNavigateUrlFormatString, dataUrlFields)
                                End If
                            End If
                        End If
                    End If

                    For Each ctrl As Control In cell.Controls
                        If Not customFieldList.Contains(ctrl.ID) Then
                            If TypeOf ctrl Is HyperLink Then
                                Dim tempURLNode As String() = CType(ctrl, HyperLink).NavigateUrl.Split("&")

                                Dim nodeStr As String()
                                Dim keyStr As String = ""
                                Dim valueStr As String = ""

                                Dim finURL As String = ""

                                For i As Integer = 0 To tempURLNode.Count - 1
                                    keyStr = ""
                                    valueStr = ""

                                    nodeStr = tempURLNode(i).Split("=")

                                    If nodeStr.Length > 1 Then
                                        keyStr = nodeStr(0)
                                        valueStr = HttpUtility.UrlEncode(nodeStr(1))

                                        If finURL <> "" Then finURL = finURL & "&"
                                        finURL = finURL & keyStr & "=" & valueStr
                                    End If
                                Next

                                If finURL <> "" Then CType(ctrl, HyperLink).NavigateUrl = finURL

                            ElseIf TypeOf ctrl Is Label Then
                                Dim colCodeName As String = ""
                                Dim fieldcol As String = ""

                                If ctrl.ID.Contains("|") Then
                                    colCodeName = ctrl.ID.Replace("|", ".")
                                    fieldcol = ctrl.ID.Split("|")(1)
                                Else
                                    colCodeName = ctrl.ID
                                    fieldcol = ctrl.ID
                                End If

                                Dim showColorYN As String = dbFun.getValueFromSQL("select srcl_show_color_yn from " &
                                                                                    "wms_search_col where fun_code = '" & menu_code & "' " &
                                                                                    "and srcl_list_seq <> 0 " &
                                                                                    "and cold_tabcol = '" & gU.dbEncode(colCodeName) & "'")

                                Dim colSQL As String = "select * from wms_col_code " &
                                                        "where colc_tabcol = '" & gU.dbEncode(colCodeName) & "' and " &
                                                        "colc_code = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim) & "'"

                                Dim colDt As New DataTable

                                colDt = gDB.getDataTable(colSQL)

                                If colDt.Rows.Count > 0 Then
                                    If Session("gLang") = "E" Then
                                        CType(ctrl, Label).Text = gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_eng_value").ToString.Trim, "")
                                    ElseIf Session("gLang") = "C" Then
                                        CType(ctrl, Label).Text = gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_chi_value").ToString.Trim, "")
                                    End If

                                    If showColorYN = "Y" AndAlso gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_color").ToString.Trim, "") <> "" Then
                                        Dim bgColor As New Drawing.Color

                                        If colDt.Rows(0).Item("colc_color").ToString.Trim.Contains("#") Then
                                            bgColor = Drawing.ColorTranslator.FromHtml(colDt.Rows(0).Item("colc_color").ToString.Trim)
                                        Else
                                            bgColor = Drawing.Color.FromName(colDt.Rows(0).Item("colc_color").ToString.Trim)
                                        End If

                                        For Each iCell As DataControlFieldCell In e.Row.Cells
                                            iCell.BackColor = bgColor
                                        Next
                                    End If
                                Else
                                    Select Case DataBinder.Eval(e.Row.DataItem, fieldcol).GetType.ToString
                                        Case "System.Decimal"
                                            CType(ctrl, Label).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, fieldcol).ToString, 0)
                                        Case "System.DBNull"
                                            CType(ctrl, Label).Text = DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim
                                        Case "System.String"
                                            CType(ctrl, Label).Text = DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim
                                        Case Else
                                            CType(ctrl, Label).Text = DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim
                                    End Select

                                End If

                            ElseIf TypeOf ctrl Is HtmlAnchor Then
                                Dim colCodeName As String = ""
                                Dim fieldcol As String = ""

                                If ctrl.ID.Contains("|") Then
                                    colCodeName = ctrl.ID.Replace("|", ".")
                                    fieldcol = ctrl.ID.Split("|")(1)
                                Else
                                    colCodeName = ctrl.ID
                                    fieldcol = ctrl.ID
                                End If

                                Dim showColorYN As String = dbFun.getValueFromSQL("select srcl_show_color_yn from " &
                                                                                    "wms_search_col where fun_code = '" & menu_code & "' " &
                                                                                    "and srcl_list_seq <> 0 " &
                                                                                    "and cold_tabcol = '" & gU.dbEncode(colCodeName) & "'")

                                Dim colSQL As String = "select * from wms_col_code " &
                                                        "where colc_tabcol = '" & gU.dbEncode(colCodeName) & "' and " &
                                                        "colc_code = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim) & "'"

                                Dim colDt As New DataTable

                                colDt = gDB.getDataTable(colSQL)

                                If colDt.Rows.Count > 0 Then
                                    If Session("gLang") = "E" Then
                                        CType(ctrl, HtmlAnchor).InnerHtml = gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_eng_value").ToString.Trim, "")
                                    ElseIf Session("gLang") = "C" Then
                                        CType(ctrl, HtmlAnchor).InnerHtml = gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_chi_value").ToString.Trim, "")
                                    End If

                                    If showColorYN = "Y" AndAlso gU.decodeNullOrEmpty(colDt.Rows(0).Item("colc_color").ToString.Trim, "") <> "" Then
                                        Dim bgColor As New Drawing.Color

                                        If colDt.Rows(0).Item("colc_color").ToString.Trim.Contains("#") Then
                                            bgColor = Drawing.ColorTranslator.FromHtml(colDt.Rows(0).Item("colc_color").ToString.Trim)
                                        Else
                                            bgColor = Drawing.Color.FromName(colDt.Rows(0).Item("colc_color").ToString.Trim)
                                        End If

                                        For Each iCell As DataControlFieldCell In e.Row.Cells
                                            iCell.BackColor = bgColor
                                        Next
                                    End If
                                Else
                                    CType(ctrl, HtmlAnchor).InnerHtml = DataBinder.Eval(e.Row.DataItem, fieldcol).ToString.Trim
                                End If

                                Dim keyCodeArray As ArrayList = ViewState("keyCodeArray")

                                If keyCodeArray IsNot Nothing AndAlso keyCodeArray.Count > 0 Then
                                    Dim args As String()

                                    For y As Integer = 0 To keyCodeArray.Count - 1
                                        ReDim Preserve args(y)

                                        args(y) = HttpUtility.UrlEncode(gU.dbEncode(DataBinder.Eval(e.Row.DataItem, keyCodeArray(y)).ToString.Trim))
                                    Next

                                    DirectCast(ctrl, HtmlAnchor).HRef = String.Format(DirectCast(ctrl, HtmlAnchor).HRef, args)
                                End If
                            ElseIf TypeOf ctrl Is CheckBox Then
                                buildAttribute(DirectCast(ctrl, CheckBox), e.Row.DataItem, DirectCast(ctrl, CheckBox).Attributes.Keys)

                                'Dim keyArray As ArrayList = ViewState("keyArray")
                                'Dim valueStr As String = ""
                                'If keyArray IsNot Nothing AndAlso keyArray.Count > 0 Then
                                '    '    Dim tempConcat As String()

                                '    '    For y As Integer = 0 To keyArray.Count - 1
                                '    '        ReDim Preserve tempConcat(y)

                                '    '        tempConcat(y) = HttpUtility.UrlEncode(gU.dbEncode(DataBinder.Eval(e.Row.DataItem, keyArray(y)).ToString.Trim))
                                '    '    Next

                                '    '    valueStr = String.Join("#", tempConcat)

                                '    '    If valueStr <> "" Then
                                '    '        Dim cbList As HiddenField = DirectCast(srchForm.FindControl("cbList"), HiddenField)

                                '    '        If Not cbList.Value.Contains(valueStr) Then
                                '    '            If cbList.Value <> "" Then cbList.Value += ","
                                '    '            cbList.Value += valueStr
                                '    '        End If
                                '    '    End If
                                'End If
                            End If
                        Else
                            If TypeOf ctrl Is HtmlAnchor Then
                                buildAttribute(DirectCast(ctrl, HtmlAnchor), e.Row.DataItem, DirectCast(ctrl, HtmlAnchor).Attributes.Keys, True)
                            ElseIf TypeOf ctrl Is HtmlInputButton Then
                                buildAttribute(DirectCast(ctrl, HtmlInputButton), e.Row.DataItem, DirectCast(ctrl, HtmlInputButton).Attributes.Keys, True)
                            ElseIf TypeOf ctrl Is CheckBox Then
                                buildAttribute(DirectCast(ctrl, CheckBox), e.Row.DataItem, DirectCast(ctrl, CheckBox).Attributes.Keys)
                            End If
                        End If
                    Next
                Next

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)
                If Not nButton Is Nothing Then
                    If Session("gLang") = "E" Then
                        nButton.Attributes.Add("onclick", "javascript:event.returnValue=confirm('Are you sure you want to delete this record?')")
                        nButton.Text = "Delete"
                    ElseIf Session("gLang") = "C" Then
                        nButton.Attributes.Add("onclick", "javascript:event.returnValue=confirm('你是否確定要刪除這個資料?')")
                        nButton.Text = "删除"
                    End If
                End If

                If srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" Then
                    Dim tempStr As String = ""
                    For i = 0 To pItemArray.Count - 1
                        Dim ptempItemArray As String()
                        ptempItemArray = Split(pItemArray(i), "|")
                        If ptempItemArray.Count > 1 Then
                            Dim sHiddenField As HiddenField = CType(e.Row.FindControl("VALUE" & ptempItemArray(1)), HiddenField)
                            If sHiddenField IsNot Nothing Then sHiddenField.Value = DataBinder.Eval(e.Row.DataItem, "VALUE" & ptempItemArray(1)).ToString.Trim
                        Else
                            Dim sHiddenField As HiddenField = CType(e.Row.FindControl("VALUE"), HiddenField)
                            If sHiddenField IsNot Nothing Then sHiddenField.Value = DataBinder.Eval(e.Row.DataItem, "VALUE").ToString.Trim
                        End If
                    Next
                Else
                    'Dim sButton As Button = CType(e.Row.FindControl("btnSelect"), Button)
                    Dim sButton As HtmlInputButton = CType(e.Row.FindControl("btnSelect"), HtmlInputButton)

                    If Not sButton Is Nothing Then
                        Dim tempStr As String = ""
                        For i = 0 To pItemArray.Count - 1
                            Dim ptempItemArray As String()
                            ptempItemArray = Split(pItemArray(i), "|")
                            If ptempItemArray.Count > 1 Then
                                If ptempItemArray.Count > 2 Then
                                    If ptempItemArray(2) = "L" Then
                                        tempStr = tempStr & "window.opener." & ptempItemArray(0) & ".innerHTML = """ & gU.jsString(DataBinder.Eval(e.Row.DataItem, "VALUE" & ptempItemArray(1)).ToString.Trim) & """;"
                                    Else
                                        tempStr = tempStr & "window.opener.document." & pForm & "." & ptempItemArray(0) & ".value = """ & gU.jsString(DataBinder.Eval(e.Row.DataItem, "VALUE" & ptempItemArray(1)).ToString.Trim) & """;"
                                    End If
                                Else
                                    tempStr = tempStr & "window.opener.document." & pForm & "." & ptempItemArray(0) & ".value = """ & gU.jsString(DataBinder.Eval(e.Row.DataItem, "VALUE" & ptempItemArray(1)).ToString.Trim) & """;"
                                End If
                            Else
                                tempStr = tempStr & "window.opener.document." & pForm & "." & pItemArray(i) & ".value = """ & gU.jsString(DataBinder.Eval(e.Row.DataItem, "VALUE").ToString.Trim) & """;"
                            End If
                        Next

                        If pFunc <> "" Then
                            tempStr = tempStr & "window.opener." & pFunc & ";"
                        End If

                        If Session("gLang") = "E" Then
                            sButton.Attributes.Add("onclick", "javascript:" & tempStr & "window.open('','_self');window.close();")
                            sButton.Value = "Select"
                        ElseIf Session("gLang") = "C" Then
                            sButton.Attributes.Add("onclick", "javascript:" & tempStr & "window.open('','_self');window.close();")
                            sButton.Value = "Select"
                        End If
                    End If
                End If
            Case DataControlRowType.Pager
                e.Row.HorizontalAlign = HorizontalAlign.Center
                e.Row.Cells(0).CssClass = "DtlLabel"
                e.Row.Cells(0).VerticalAlign = VerticalAlign.Middle
        End Select
    End Sub

    Protected Sub MasterList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_MasterGridView"), GridView)

        gvrsList.PageIndex = e.NewPageIndex

        Session("SEARCH_SESSION_PAGE_MasterGV_PAGEINDEX") = e.NewPageIndex
        Call BindGV(True)
    End Sub


    Protected Sub gvrsList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)

                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkBox.Checked Then
                    PersistRowIndex(container.DataItemIndex)
                Else
                    RemoveRowIndex(container.DataItemIndex)
                End If
            Next
        End If

        gvrsList.PageIndex = e.NewPageIndex

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = e.NewPageIndex

        Call BindGV()
    End Sub

    Protected Sub gvrsList_PagerIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim ddlPager As DropDownList = DirectCast(gvrsList.BottomPagerRow.FindControl("pager_select"), DropDownList)
        Dim PageNo As Integer = Convert.ToInt32(ddlPager.SelectedItem.Value)

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)

                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkBox.Checked Then
                    PersistRowIndex(container.DataItemIndex)
                Else
                    RemoveRowIndex(container.DataItemIndex)
                End If
            Next
        End If

        gvrsList.PageIndex = PageNo - 1

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = PageNo - 1

        Call BindGV()

    End Sub

    Protected Sub gvrsList_PagerIndexForward(ByVal sender As Object, ByVal e As EventArgs)
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim ddlPager As DropDownList = DirectCast(gvrsList.BottomPagerRow.FindControl("pager_select"), DropDownList)

        Dim PageIdx As Integer = gvrsList.PageIndex + 1

        If PageIdx > gvrsList.PageCount - 1 Then
            PageIdx = gvrsList.PageCount - 1
        End If

        Dim PageNo As Integer = PageIdx + 1

        If PageNo > gvrsList.PageCount Then
            PageNo = gvrsList.PageCount
            PageIdx = gvrsList.PageCount - 1
        End If

        ddlPager.SelectedValue = PageNo

        Dim checkedAll As Boolean = False

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)
                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkAll.Checked Then
                    checkedAll = True
                    If chkBox.Checked = False Then
                        chkBox.Checked = True
                    End If
                Else
                    If chkBox.Checked Then
                        PersistRowIndex(container.DataItemIndex)
                    Else
                        RemoveRowIndex(container.DataItemIndex)
                    End If
                End If
            Next
        End If

        gvrsList.PageIndex = PageIdx

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = PageIdx

        Call BindGV()

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkNextPageAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            If checkedAll Then chkNextPageAll.Checked = True

            RePopulateCheckBoxes()
        End If
    End Sub

    Protected Sub gvrsList_PagerIndexBackward(ByVal sender As Object, ByVal e As EventArgs)

        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim ddlPager As DropDownList = DirectCast(gvrsList.BottomPagerRow.FindControl("pager_select"), DropDownList)

        Dim PageIdx As Integer = gvrsList.PageIndex - 1

        If PageIdx < 0 Then
            PageIdx = 0
        End If

        Dim PageNo As Integer = PageIdx + 1

        ddlPager.SelectedValue = PageNo

        Dim checkedAll As Boolean = False

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)

                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkAll.Checked Then
                    checkedAll = True
                    If chkBox.Checked = False Then
                        chkBox.Checked = True
                    End If
                Else
                    If chkBox.Checked Then
                        PersistRowIndex(container.DataItemIndex)
                    Else
                        RemoveRowIndex(container.DataItemIndex)
                    End If
                End If
            Next
        End If

        gvrsList.PageIndex = PageIdx

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = PageIdx

        Call BindGV()

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkNextPageAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            If checkedAll Then chkNextPageAll.Checked = True

            RePopulateCheckBoxes()
        End If
    End Sub

    Protected Sub gvrsList_PagerIndexFirst(ByVal sender As Object, ByVal e As EventArgs)

        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim ddlPager As DropDownList = DirectCast(gvrsList.BottomPagerRow.FindControl("pager_select"), DropDownList)

        Dim PageIdx As Integer = 0

        Dim PageNo As Integer = PageIdx + 1

        ddlPager.SelectedValue = PageNo

        Dim checkedAll As Boolean = False

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)

                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkAll.Checked Then
                    checkedAll = True
                    If chkBox.Checked = False Then
                        chkBox.Checked = True
                    End If
                Else
                    If chkBox.Checked Then
                        PersistRowIndex(container.DataItemIndex)
                    Else
                        RemoveRowIndex(container.DataItemIndex)
                    End If
                End If
            Next
        End If

        gvrsList.PageIndex = PageIdx

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = PageIdx

        Call BindGV()

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkNextPageAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            If checkedAll Then chkNextPageAll.Checked = True

            RePopulateCheckBoxes()
        End If
    End Sub

    Protected Sub gvrsList_PagerIndexLast(ByVal sender As Object, ByVal e As EventArgs)

        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        Dim ddlPager As DropDownList = DirectCast(gvrsList.BottomPagerRow.FindControl("pager_select"), DropDownList)

        Dim PageIdx As Integer = gvrsList.PageCount - 1

        Dim PageNo As Integer = PageIdx + 1

        ddlPager.SelectedValue = PageNo

        Dim checkedAll As Boolean = False

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)

                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)

                If chkAll.Checked Then
                    checkedAll = True
                    If chkBox.Checked = False Then
                        chkBox.Checked = True
                    End If
                Else
                    If chkBox.Checked Then
                        PersistRowIndex(container.DataItemIndex)
                    Else
                        RemoveRowIndex(container.DataItemIndex)
                    End If
                End If
            Next
        End If

        gvrsList.PageIndex = PageIdx

        Session("SEARCH_SESSION_PAGE_GV_PAGEINDEX") = PageIdx

        Call BindGV()

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim chkNextPageAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            If checkedAll Then chkNextPageAll.Checked = True

            RePopulateCheckBoxes()
        End If
    End Sub

    Protected Sub gvrsList_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        Dim dt As DataTable = TryCast(ViewState("dt"), DataTable)

        If dt IsNot Nothing Then
            Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)

            'Sort the data.
            dt.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
            gvrsList.DataSource = dt
            gvrsList.DataBind()

        End If
    End Sub

    Private Sub search_fun()
        Session.Remove("SEARCH_SESSION_PAGE_CRITERIA_LIST")
        ViewState.Remove("SEARCH_VIEWSTATE_PAGE_SelectedCBIndex")

        If ViewState("noresetbtn") <> "Y" Then
            Session("SEARCH_SESSION_PAGE_REMAIN_STATUS") = True
            Session("SEARCH_SESSION_PAGE_STATUS_MENU_CODE") = menu_code
        End If

        Session("SEARCH_SESSION_PAGE_SEARCHED") = "Y"

        'load_ModalPopupExtender.Show()

        If (srch_table.Rows(0).Item("SRCH_MULTIPLE_LOOKUP").ToString.Trim.ToUpper = "Y" OrElse
            srch_table.Rows(0).Item("SRCH_ITEM_SELECTION_YN").ToString.Trim.ToUpper = "Y") Then
            Dim cbSelected As HiddenField = TryCast(srchForm.FindControl("cbSelected"), HiddenField)
            Dim cbUnSelected As HiddenField = TryCast(srchForm.FindControl("cbUnSelected"), HiddenField)
            Dim toggle_selectAll As HiddenField = TryCast(srchForm.FindControl("toggle_selectAll"), HiddenField)

            If cbSelected IsNot Nothing Then cbSelected.Value = ""
            If cbUnSelected IsNot Nothing Then cbUnSelected.Value = ""
            If toggle_selectAll IsNot Nothing Then toggle_selectAll.Value = "N"
        End If

        If srch_table.Rows(0).Item("SRCH_DIRECTLY_REDIRECT_YN").ToString.ToUpper.Trim = "Y" Then

            Dim redirectURL As String = srch_table.Rows(0).Item("SRCH_REDIRECT_URL").ToString.Trim
            Dim target As String = srch_table.Rows(0).Item("SRCH_REDIRECT_OPEN_TARGET").ToString.Trim

            Dim sSQL As String = ""
            Dim wSQL As String = ""
            Dim cSQL As String = ""
            Dim aSQL As String = ""
            Dim oSQL As String = ""
            Dim fieldList As String = ""

            Dim ctrlArray As New Dictionary(Of String, String())
            Dim abortFlag As Boolean = False
            Dim fPost As New FormPosting

            Dim theSQL As String = generateSQL(False, , , fPost, sSQL, wSQL, cSQL, aSQL, oSQL, fieldList, ctrlArray, abortFlag)

            If abortFlag Then Exit Sub

            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                GenericAppMod.ctrlValue = ctrlArray
                theSQL = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, sSQL, wSQL, cSQL, aSQL, oSQL, fieldList)
            End If

            Session("GENERIC_SESSION_SRCH_COMPLETE_SQL") = theSQL
            Session("GENERIC_SESSION_SRCH_SQL") = sSQL
            Session("GENERIC_SESSION_SRCH_WHERE_SQL") = wSQL
            Session("GENERIC_SESSION_SRCH_CRITERIA_SQL") = cSQL
            Session("GENERIC_SESSION_SRCH_ADD_SQL") = aSQL
            Session("GENERIC_SESSION_SRCH_ORDERBY_SQL") = oSQL
            Session("GENERIC_SESSION_SRCH_SELECT_FIELDS") = fieldList

            If redirectURL <> "" Then
                'fPost.Url = redirectURL
                'If target <> "" Then fPost.FormTarget = target
                'fPost.Add("GENERIC_REQUEST_COMPLETE_SQL", theSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_SQL", sSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_WHERE_SQL", wSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_CRITERIA_SQL", cSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_ADD_SQL", aSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_ORDERBY_SQL", oSQL)
                'fPost.Add("GENERIC_REQUEST_SRCH_SELECT_FIELDS", fieldList)

                'fPost.Post()

                Dim rmtPost As New RemotePost

                rmtPost.Url = redirectURL
                If target <> "" Then rmtPost.Target = target
                rmtPost.Add("GENERIC_REQUEST_COMPLETE_SQL", theSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_SQL", sSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_WHERE_SQL", wSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_CRITERIA_SQL", cSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_ADD_SQL", aSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_ORDERBY_SQL", oSQL)
                rmtPost.Add("GENERIC_REQUEST_SRCH_SELECT_FIELDS", fieldList)
                rmtPost.Post()
            End If
        Else
            Dim oGridView As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)

            If oGridView.AllowPaging = True Then
                If oGridView.Rows.Count > 0 Then
                    Dim ddlPager As DropDownList = TryCast(oGridView.BottomPagerRow.FindControl("pager_select"), DropDownList)

                    If ddlPager IsNot Nothing Then
                        ddlPager.SelectedIndex = 0
                    End If
                End If
            End If

            If srch_table.Rows(0).Item("SRCH_LOOKUP_MC_YN").ToString.ToUpper.Trim = "Y" Then
                'Call GenerateMasterGV()
                Call BindGV(True, , True)
            Else
                Call BindGV(, , True)
            End If
        End If
    End Sub

    Protected Sub click_search(ByVal sender As Object, ByVal e As System.EventArgs)
        search_fun()
    End Sub

    Protected Sub click_reset(ByVal sender As Object, ByVal e As System.EventArgs)
        appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")

        Try

            Response.Redirect(Request.Url.ToString, False)

        Catch ex As Exception

        End Try
    End Sub

    'For SHORTSHIP REPORT
    Protected Sub click_SSreport(ByVal sender As Object, ByVal e As System.EventArgs)
        showSSReport()
    End Sub

    'For PICKLIST REPORT
    Protected Sub click_report(ByVal sender As Object, ByVal e As System.EventArgs)
        showDtlReport()
    End Sub

    'For STK Excel REPORT
    Protected Sub click_Exp_STKReport(ByVal sender As Object, ByVal e As System.EventArgs)
        showExpSTKReport()
    End Sub

    Protected Sub click_download(ByVal sender As Object, ByVal e As System.EventArgs)
        showReport()
    End Sub

    Protected Sub click_print(ByVal sender As Object, ByVal e As System.EventArgs)
        showReport(True)
    End Sub

    Protected Sub cbSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tempStr As String = ""
        For i = 0 To newDataArray.Count - 1
            newDataArray(i) = ""
        Next
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        For i1 As Integer = 0 To gvrsList.Rows.Count - 1
            Dim sCheckBox As CheckBox = CType(gvrsList.Rows(i1).FindControl("btnCB"), CheckBox)
            If sCheckBox.Checked Then
                For i = 0 To pItemArray.Count - 1
                    Dim ptempItemArray As String()
                    ptempItemArray = Split(pItemArray(i), "|")

                    If newDataArray(i) <> "" Then newDataArray(i) += ", "

                    If ptempItemArray.Count > 1 Then
                        Dim sHiddenField As HiddenField = CType(gvrsList.Rows(i1).FindControl("VALUE" & ptempItemArray(1)), HiddenField)

                        If sHiddenField IsNot Nothing Then
                            newDataArray(i) += sHiddenField.Value.ToString
                        Else
                            If ptempItemArray(1).Contains("#") Then
                                Dim ccStr As String() = Split(ptempItemArray(1), "#")

                                If ccStr.Count > 1 Then
                                    If gvrsList.Rows(i1).FindControl(ccStr(0)) IsNot Nothing Then
                                        If gvrsList.Rows(i1).FindControl(ccStr(0)).FindControl(ccStr(1)) IsNot Nothing Then
                                            Dim ctl As Control = gvrsList.Rows(i1).FindControl(ccStr(0)).FindControl(ccStr(1))
                                            Select Case TypeName(ctl).ToUpper
                                                Case "TEXTBOX"
                                                    newDataArray(i) += DirectCast(ctl, TextBox).Text

                                                Case "DROPDOWNLIST"
                                                    newDataArray(i) += DirectCast(ctl, DropDownList).SelectedValue

                                            End Select
                                        Else
                                            Dim ctl As Control = gvrsList.Rows(i1).FindControl(ccStr(0))

                                            Select Case TypeName(ctl).ToUpper
                                                Case "TEXTBOX"
                                                    newDataArray(i) += DirectCast(ctl, TextBox).Text

                                                Case "DROPDOWNLIST"
                                                    newDataArray(i) += DirectCast(ctl, DropDownList).SelectedValue

                                            End Select
                                        End If
                                    End If
                                Else
                                    If gvrsList.Rows(i1).FindControl(ccStr(0)) IsNot Nothing Then
                                        Dim ctl As Control = gvrsList.Rows(i1).FindControl(ccStr(0))

                                        Select Case TypeName(ctl).ToUpper
                                            Case "TEXTBOX"
                                                newDataArray(i) += DirectCast(ctl, TextBox).Text

                                            Case "DROPDOWNLIST"
                                                newDataArray(i) += DirectCast(ctl, DropDownList).SelectedValue

                                        End Select
                                    End If
                                End If
                            Else
                                If gvrsList.Rows(i1).FindControl(ptempItemArray(1)) IsNot Nothing Then
                                    Dim ctl As Control = gvrsList.Rows(i1).FindControl(ptempItemArray(1))

                                    Select Case TypeName(ctl).ToUpper
                                        Case "TEXTBOX"
                                            newDataArray(i) += DirectCast(ctl, TextBox).Text

                                        Case "DROPDOWNLIST"
                                            newDataArray(i) += DirectCast(ctl, DropDownList).SelectedValue

                                    End Select
                                End If
                            End If
                        End If
                    Else
                        Dim sHiddenField As HiddenField = CType(gvrsList.Rows(i1).FindControl("VALUE"), HiddenField)

                        If sHiddenField IsNot Nothing Then
                            newDataArray(i) += sHiddenField.Value.ToString
                        End If
                    End If
                Next
            End If
        Next

        For i = 0 To pItemArray.Count - 1
            Dim ptempItemArray As String()
            ptempItemArray = Split(pItemArray(i), "|")
            If ptempItemArray.Count > 1 Then
                tempStr = tempStr & "window.opener.document." & pForm & "." & ptempItemArray(0) & ".value = """ & HttpUtility.HtmlEncode(newDataArray(i).ToString) & """;"
            Else
                tempStr = tempStr & "window.opener.document." & pForm & "." & pItemArray(i) & ".value = """ & HttpUtility.HtmlEncode(newDataArray(i).ToString) & """;"
            End If
        Next

        Dim client1 As ClientScriptManager
        client1 = Me.ClientScript
        Dim strScript As String = ""
        If pFunc <> "" Then
            strScript = tempStr & "window.opener." & pFunc & ";window.open('','_self');window.close();"
        Else
            strScript = tempStr & "window.open('','_self');window.close();"
        End If
        'If (Not client1.IsStartupScriptRegistered(Me.GetType(), "")) Then
        '    client1.RegisterStartupScript(Me.GetType(), "", strScript, True)
        'End If

        System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "lookupSubmit", strScript, True)
    End Sub

    Protected Sub new_redirect(ByVal sender As Object, ByVal e As System.EventArgs)
        If ViewState("noresetbtn") <> "Y" Then
            Session("SEARCH_SESSION_PAGE_REMAIN_STATUS") = True
            Session("SEARCH_SESSION_PAGE_STATUS_MENU_CODE") = menu_code
        End If

        Dim SCString As String = ""
        Dim symbol As String = srch_table.Rows(0).Item("srch_app_sql").ToString

        For x As Integer = 0 To col_table.Rows.Count - 1
            Dim FieldName As String = "'"
            Dim ctrlName As String = ""
            Dim oper As String = ""

            If col_table.Rows(x).Item("srcl_custom_field_yn").ToString.Trim.ToUpper = "Y" Then
                Continue For
            End If

            FieldName = col_table.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper
            ctrlName = FieldName.Replace(".", "_").Trim
            oper = col_table.Rows(x).Item("srcl_operator").ToString.Trim.ToUpper


            If col_table.Rows(x).Item("SRCL_ISREADONLY_YN").ToString.Trim = "Y" Then
                If col_table.Rows(x).Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.Trim = "" Then
                    GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString, oper, ctrlName, FieldName, symbol)
                Else
                    If col_table.Rows(x).Item("srcl_edit_style").ToString = "S" Then
                        GenerateObjectSQL(SCString, col_table.Rows(x), "H", oper, ctrlName, FieldName, symbol, "S")

                        GenerateObjectSQL(SCString, col_table.Rows(x), "L", oper, ctrlName, FieldName, symbol, col_table.Rows(x).Item("srcl_edit_style").ToString.Trim)
                    End If
                End If
            Else
                GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString.Trim, oper, ctrlName, FieldName, symbol)
            End If

            'If col_table.Rows(x).Item("srcl_add_where_clause").ToString.ToUpper.Trim <> "" Then
            '    SCString = SCString & col_table.Rows(x).Item("srcl_add_where_clause").ToString.ToUpper.Trim
            'End If
        Next

        Dim urlTar As String = srch_table.Rows(0).Item("srch_redirect_url").ToString

        If urlTar.Trim.Contains("?") Then
            If urlTar.Trim.Substring(urlTar.Trim.Length - 1, 1) = "&" Then
                Response.Redirect(urlTar & "mode=N&menu_code=" & menu_code)
            Else
                Response.Redirect(urlTar & "&mode=N&menu_code=" & menu_code)
            End If
        Else
            Response.Redirect(urlTar & "?mode=N&menu_code=" & menu_code)
        End If

        Response.End()
    End Sub

    'for Picklist report'
    Private Sub showDtlReport()
        Dim sqlString As String
        Dim srchSQL As String = ""
        Dim whereSQL As String = ""
        Dim criteriaSQL As String = ""
        Dim addSQL As String = ""
        Dim orderbySQL As String = ""
        Dim FieldList As String = ""
        Dim masterWhereString As String = ""
        Dim ctrlArray As New Dictionary(Of String, String())
        Dim abortFlag As Boolean = False
        Dim isMaster As Boolean = False
        Dim nDataSource As DataTable
        'sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
        If masterWhereString <> "" Then
            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                GenericAppMod.ctrlValue = ctrlArray
                sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
            Else
                sqlString = generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
            End If
        Else
            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                GenericAppMod.ctrlValue = ctrlArray
                sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
            Else
                sqlString = generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
            End If
        End If

        nDataSource = gDB.getDataTable(sqlString)

        Dim DOList As New List(Of String)
        For index As Integer = 0 To nDataSource.Rows.Count - 1
            Dim do_code = nDataSource.Rows(index)("DO_CODE").ToString
            DOList.Add(do_code.Trim)
        Next

        Session("do_code_list") = DOList.Distinct().ToList()
        Dim url As String = "~/../OUTBOUND/DO/DO_DETAIL/dodtl_list.aspx"
        Dim s As String = "window.open('" & url + "', 'DO_DETAIL', 'width=800,height=800,left=100,top=100,resizable=yes');"
        ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)

        'OLD CODE'
        'Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        'Dim DOList As New List(Of String)
        'For Each gvr As GridViewRow In gvrsList.Rows
        '    'Dim cb As CheckBox = DirectCast(gvr.FindControl("btnCB"), CheckBox)
        '    Dim do_code = DirectCast(gvr.Cells(0).Controls(0), HyperLink).Text
        '    DOList.Add(do_code.Trim)
        'Next
        'Session("do_code_list") = DOList.Distinct().ToList()
        'Dim url As String =  "~/../OUTBOUND/DO/DO_DETAIL/dodtl_list.aspx"
        'Dim s As String = "window.open('" & url + "', 'DO_DETAIL', 'width=800,height=800,left=100,top=100,resizable=yes');"
        'ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)

    End Sub

    'For SHORTSHIP REPORT
    Private Sub showSSReport()
        Dim sqlString As String
        Dim srchSQL As String = ""
        Dim whereSQL As String = ""
        Dim criteriaSQL As String = ""
        Dim addSQL As String = ""
        Dim orderbySQL As String = ""
        Dim FieldList As String = ""
        Dim masterWhereString As String = ""
        Dim ctrlArray As New Dictionary(Of String, String())
        Dim abortFlag As Boolean = False
        Dim isMaster As Boolean = False
        Dim nDataSource As DataTable
        'sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)

        If masterWhereString <> "" Then
            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                GenericAppMod.ctrlValue = ctrlArray
                sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
            Else
                sqlString = generateSQL(isMaster, masterWhereString, , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
            End If
        Else
            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, ctrlArray, abortFlag)

                GenericAppMod.ctrlValue = ctrlArray
                sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList)
            Else
                sqlString = generateSQL(isMaster, , , , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, FieldList, , abortFlag)
            End If
        End If

        nDataSource = gDB.getDataTable(sqlString)

        Dim DOList As New List(Of String)
        For index As Integer = 0 To nDataSource.Rows.Count - 1
            Dim do_code = nDataSource.Rows(index)("DO_CODE").ToString
            DOList.Add(do_code.Trim)
        Next

        Session("do_code_list") = DOList.Distinct().ToList()
        Dim url As String = "~/../OUTBOUND/DO/SHORT_SHIP/short_ship.aspx"
        Dim s As String = "window.open('" & url + "', 'DO_PICKLIST_D', 'width=800,height=800,left=100,top=100,resizable=yes');"
        ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)

        'OLD CODE'
        'Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)
        'Dim DOList As New List(Of String)
        'For Each gvr As GridViewRow In gvrsList.Rows
        '    'Dim cb As CheckBox = DirectCast(gvr.FindControl("btnCB"), CheckBox)
        '    Dim do_code = DirectCast(gvr.Cells(0).Controls(0), HyperLink).Text
        '    DOList.Add(do_code.Trim)

        'Next
        'Session("do_code_list") = DOList.Distinct().ToList()
        'Dim url As String = "~/../OUTBOUND/DO/SHORT_SHIP/short_ship.aspx"
        'Dim s As String = "window.open('" & url + "', 'DO_PICKLIST_D', 'width=800,height=800,left=100,top=100,resizable=yes');"
        'ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)

    End Sub

    'For STK Excel REPORT
    Private Sub showExpSTKReport()
        Dim sqlString As String
        Dim nDataSource As DataSet

        Dim batchSQL As String = "WMS_ITEM_LOC_BAL.ILOC_BATCH_NO"
        Try

            sqlString = "SELECT WMS_STORER.STO_SHORTNAME,  WMS_ITEM.ITM_SKU_NO, WMS_WAREHOUSE.WH_MAIN_WH, " &
                                    "WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, " &
                                    "WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY, " &
                                    "ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) as ITM_PCS_PER_UOM, (ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) * WMS_ITEM_LOC_BAL.ILOC_BAL_QTY ) as TOTAL_NO, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_CBM, WMS_ITEM_LOC_BAL.ILOC_BAL_KG, " &
                                    batchSQL & ", " &
                                    "Convert(varchar, ILOC_EXPIRY_DATE, 103) AS ILOC_EXPIRY_DATE, " &
                                    "Convert(varchar, ILOC_MANU_DATE, 103) AS ILOC_MANU_DATE, " &
                                    "WMS_ITEM_LOC_BAL.VND_CODE, " &
                                    "WMS_item.itm_name, WMS_ITEM.ITM_TEMP_FR, WMS_ITEM.ITM_TEMP_TO, " &
                                    "ISNULL(WMS_WAREHOUSE.WH_NAME, WMS_ITEM_LOC_BAL.ILOC_WH) as WH_NAME, " &
                                    "ISNULL(WMS_WH_FL.FL_NAME, WMS_ITEM_LOC_BAL.ILOC_FLOOR) as FL_NAME, " &
                                    "ISNULL(WMS_WH_AREA.AR_NAME, WMS_ITEM_LOC_BAL.ILOC_AREA) as AR_NAME, " &
                                    "ISNULL(WMS_WH_RACK.RK_NAME, WMS_ITEM_LOC_BAL.ILOC_RACK) as RK_NAME, " &
                                    "TOTAL_BAL.TOTAL_BAL_QTY, HOLD_STOCK.HOLD_QTY, " &
                                    "PICK_ITEM.PICKED_QTY, TOTAL_PICK.TOTAL_PICKED_QTY, CO_ITEM.total_co_qty, " &
                                    "ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) as STOCK_AVAIL_QTY, " &
                                    "ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) - ISNULL(TOTAL_PICK.TOTAL_PICKED_QTY, 0) as AVAIL_QTY, " &
                                    "t.aitm_qty_per_ctn, " &
                                    "case when t.aitm_qty_per_ctn > 0 then ceiling(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) else 0 end as NO_OF_CARTON, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.carton_cbm as TOTAL_carton_cbm, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.AITM_VOL as TOTAL_carton_KG, " &
                                    "'' as ILBS_SERIAL_NO, NULL as ILBS_QTY2, '' as ILBS_DRUM_ID, " &
                                    " WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_EMB, " &
                                    " CONVERT(varchar, WMS_ITEM.ITM_PD_RCV_DATE, 103) as ITM_PD_RCV_DATE, WMS_ITEM.ITM_PD_CONTRACT_NO, WMS_ITEM.ITM_PD_ARR_NOTICE_NO, WMS_ITEM.ITM_PD_COND_OF_SPARES, WMS_ITEM.ITM_PD_ST_1_YEAR, WMS_ITEM.ITM_PD_ST_OVER_1_YEAR, " &
                                    " WMS_ITEM.ITM_DRAWING_NO, WMS_ITEM.ITM_SERIAL_NO_YN, WMS_ITEM.ITM_STACKABLE_YN, WMS_ITEM.ITM_INSP_YN, WMS_ITEM.ITM_SCRAP_YN,  " &
                                    " WMS_ITEM.ITM_NONSTOCK_YN, WMS_ITEM.ITM_DG_YN, WMS_ITEM.ITM_REQ_STORE_HUM_YN, WMS_ITEM.ITM_REQ_STORE_AIRCON_YN,  " &
                                    " WMS_ITEM.ITM_CHE_CLASS, WMS_ITEM.ITM_REQ_MSDS_YN, WMS_ITEM.ITM_WEIGHT_TYPE, WMS_ITEM.ITM_ORO_YN, WMS_ITEM.ITEM_PRICE_CLASS,  " &
                                    " WMS_ITEM_LOC_BAL.SYS_LUB, WMS_ITEM_LOC_BAL.SYS_LUD,WMS_ITEM_LOC_BAL.SYS_LUD, WMS_ITEM.ITM_ROP_APL, WMS_ITEM.ITM_ROP_LAM, WMS_ITEM.ITM_ROP_CABLE,WMS_ITEM.ITM_ROP_NP, " &
                                    " replace(replace(replace(replace(replace(replace(replace(replace(replace(WMS_ITEM.ITM_CHE_CLASS,'1F','Flammable'),'2E','Explosive'),'3O', 'Oxidizing'),'4H','Harmful'),'5T','Toxic'),'6C','Corrosive'),'7I','Irritant'),'8C','Carcinogen'),'NA','NA') as che_class, " &
                                    " gr_itm.total_rcv_qty,WMS_ITEM_LOC_BAL.ILOC_PO_NO, NUll as BORD_QTY, WMS_ITEM.ITM_CRITICAL_YN, WMS_ITEM.ITM_RESTRICTED_ITEM,'' as ROQ, " &
                                    " case when datediff(d,convert(date,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE),convert(date, getdate())) > 3 then 'Y' else 'N' end as IS_OVERDUE, " &
                                    " insp.gri_insp_qty, WMS_ITEM.ITM_UOM, '' AS ILBS_UOM2 " &
                                "FROM WMS_ITEM_LOC_BAL " &
                                "INNER JOIN WMS_STORER " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_STORER.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=WMS_STORER.STORER_CODE " &
                                "LEFT OUTER JOIN WMS_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=WMS_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=WMS_ITEM.ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=WMS_ITEM.PACK_KEY " &
                                "LEFT OUTER JOIN V_ALT_VEND_ITEM t " &
                                    "ON WMS_ITEM.IMP_CODE = t.IMP_CODE " &
                                    "AND WMS_ITEM.STORER_CODE = t.STORER_CODE " &
                                    "AND WMS_ITEM.ITM_CODE = t.ITM_CODE " &
                                    "AND WMS_ITEM.PACK_KEY = t.PACK_KEY " &
                                "LEFT OUTER JOIN WMS_WAREHOUSE " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WAREHOUSE.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WAREHOUSE.WH_CODE " &
                                "LEFT OUTER JOIN WMS_WH_FL " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_FL.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_FL.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_FL.FL_NUM " &
                                "LEFT OUTER JOIN WMS_WH_AREA " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_AREA.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_AREA.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_AREA.FL_NUM " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_AREA.AR_CODE " &
                                "LEFT OUTER JOIN WMS_WH_RACK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_RACK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_RACK.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_RACK.FL_NUM " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_RACK.AR_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_RACK=WMS_WH_RACK.RK_CODE " &
                                "LEFT OUTER JOIN ( " &
                                        "select imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                                        "ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, sum(ILOC_BAL_QTY) as TOTAL_BAL_QTY " &
                                        "from WMS_ITEM_LOC_BAL " &
                                        "where ISNULL(ILOC_BAL_QTY, 0) != 0 " &
                                        "group by imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, '')) TOTAL_BAL " &
                                    "on WMS_ITEM_LOC_BAL.IMP_CODE=TOTAL_BAL.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=TOTAL_BAL.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=TOTAL_BAL.ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=TOTAL_BAL.PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=TOTAL_BAL.ILOC_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')=TOTAL_BAL.ILOC_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select hold.imp_code, " &
                                            "hold.storer_code, " &
                                            "hold.COD_ITM_CODE, " &
                                            "hold.COD_PACK_KEY, " &
                                            "hold.COD_PALLET_NO, " &
                                            "hold.COD_BATCH_NO, " &
                                            "sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " &
                                        "from ( " &
                                            "select h.imp_code, " &
                                                "h.storer_code, " &
                                                "h.co_code, " &
                                                "h.COH_ITM_CODE as COD_ITM_CODE, " &
                                                "h.COH_PACK_KEY as COD_PACK_KEY, " &
                                                "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                                                "ISNULL(h.COH_BATCH_NO, '') as COD_BATCH_NO, " &
                                                "sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
                                            "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                            "where h.imp_code = d.imp_code " &
                                            "and h.storer_code = d.storer_code " &
                                            "and h.co_code = d.co_code " &
                                            "and h.cod_seq = d.cod_seq " &
                                            "and c.imp_code = d.imp_code " &
                                            "and c.storer_code = d.storer_code " &
                                            "and c.co_code = d.co_code " &
                                            "and h.coh_status <> 'RELEASE' " &
                                            "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                                            "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " &
                                            "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), ISNULL(h.COH_BATCH_NO, '')) hold " &
                                        "left outer join " &
                                            "( " &
                                            "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " &
                                            "from wms_do_picklist_d p, wms_delv_order d " &
                                            "where d.imp_code = p.imp_code " &
                                            "and d.storer_code = p.storer_code " &
                                            "and d.do_code = p.do_code " &
                                            "and d.do_status = 'PICKED' " &
                                            "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " &
                                            ") pick " &
                                        "on hold.imp_code = pick.imp_code " &
                                        "and hold.storer_code = pick.storer_code " &
                                        "and hold.co_code = pick.do_co_code " &
                                        "and hold.COD_ITM_CODE = pick.pld_item_no " &
                                        "and hold.COD_PACK_KEY = pick.pld_pack_key " &
                                        "and hold.COD_PALLET_NO = pick.pld_pallet_no " &
                                        "and hold.COD_BATCH_NO = pick.pld_batch_no " &
                                        "group by hold.imp_code, hold.storer_code, hold.COD_ITM_CODE, hold.COD_PACK_KEY, hold.COD_PALLET_NO, hold.COD_BATCH_NO) HOLD_STOCK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=HOLD_STOCK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=HOLD_STOCK.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=HOLD_STOCK.COD_ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=HOLD_STOCK.COD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=HOLD_STOCK.COD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')=HOLD_STOCK.COD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                                        "from wms_do_picklist_d p, wms_delv_order d " &
                                        "where d.imp_code = p.imp_code " &
                                        "and d.storer_code = p.storer_code " &
                                        "and d.do_code = p.do_code " &
                                        "and d.do_status = 'PICKED' " &
                                        "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = PICK_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = PICK_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_LOC = PICK_ITEM.PLD_LOC " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as total_picked_qty " &
                                        "from wms_do_picklist_d p, wms_delv_order d " &
                                        "where d.imp_code = p.imp_code " &
                                        "and d.storer_code = p.storer_code " &
                                        "and d.do_code = p.do_code " &
                                        "and d.do_status = 'PICKED' " &
                                        "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '')) TOTAL_PICK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = TOTAL_PICK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = TOTAL_PICK.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = TOTAL_PICK.PLD_ITEM_NO " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = TOTAL_PICK.PLD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = TOTAL_PICK.PLD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = TOTAL_PICK.PLD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no, sum(cod_os_qty) as total_co_qty " &
                                        "from ( " &
                                            "select c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000') as cod_pallet_no, ISNULL(c2.cod_batch_no, '') as cod_batch_no, " &
                                            "ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) as cod_os_qty " &
                                            "from WMS_CUST_ORDER_D c2, WMS_CUST_ORDER c1 " &
                                            "where c1.imp_code = c2.imp_code " &
                                            "and c1.storer_code = c2.storer_code " &
                                            "and c1.co_code = c2.co_code " &
                                            "and c1.co_status not in ('CANCELLED', 'CLOSED') " &
                                            "group by c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000'), ISNULL(c2.cod_batch_no, '') " &
                                            "having ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) > 0) tbCO " &
                                        "group by imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no) CO_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = CO_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = CO_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = CO_ITEM.COD_ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = CO_ITEM.COD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = CO_ITEM.COD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = CO_ITEM.COD_BATCH_NO " &
                                " LEFT OUTER JOIN (" &
                                    " SELECT WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE, ISNULL(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') AS GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, '') AS GRD_BATCH_NO, " &
                                        " WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY, " &
                                        " SUM(WMS_GOODSRCV_D.GRD_RCV_QTY) AS TOTAL_RCV_QTY " &
                                    " FROM WMS_GOODSRCV " &
                                    " INNER JOIN WMS_GOODSRCV_D " &
                                    " ON WMS_GOODSRCV.IMP_CODE = WMS_GOODSRCV_D.IMP_CODE AND  " &
                                        " WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE AND " &
                                        " WMS_GOODSRCV.GR_CODE = WMS_GOODSRCV_D.GR_CODE " &
                                    " WHERE WMS_GOODSRCV.GR_STATUS NOT IN ('CANCELLED','POSTED') " &
                                    " GROUP BY WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE, WMS_GOODSRCV_D.GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, ''), " &
                                            " WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY) GR_ITM  " &
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = GR_ITM.IMP_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.STORER_CODE = GR_ITM.STORER_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.ITM_CODE = GR_ITM.GRD_ITM_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.PACK_KEY = GR_ITM.GRD_PACK_KEY AND " &
                                    " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = GR_ITM.GRD_PALLET_NO AND " &
                                    " ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = GR_ITM.GRD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000') as gra_pallet_no, isnull(p.gra_batch_no, '') as gra_batch_no, p.gra_loc, sum(i.gri_insp_qty) as gri_insp_qty " &
                                        "from wms_goodsrcv_pa p, wms_goodsrcv_insp i " &
                                        "where p.imp_code = i.imp_code " &
                                        "and p.storer_code = i.storer_code " &
                                        "and p.gr_code = i.gr_code " &
                                        "and p.gra_itm_code = i.gri_itm_code " &
                                        "and p.gra_pack_key = i.gri_pack_key " &
                                        "and p.gra_pallet_no = i.gri_pallet_no " &
                                        "and p.gra_batch_no = i.gri_batch_no " &
                                        "and i.gri_status <> 'DONE' " &
                                        "group by p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000'), isnull(p.gra_batch_no, ''), p.gra_loc) insp " &
                                    "on WMS_ITEM_LOC_BAL.IMP_CODE = insp.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = insp.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = insp.gra_itm_code " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = insp.gra_pack_key " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = insp.gra_pallet_no " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = insp.gra_batch_no " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_LOC, '') = insp.gra_loc " &
                                "WHERE 2=2 AND WMS_ITEM_LOC_BAL.ILOC_BAL_QTY != 0 and ISNULL(wms_item.ITM_SERIAL_NO_YN,'N') <> 'Y' "


            nDataSource = gDB.getDataSet(sqlString)

            WriteXLSFile(nDataSource)

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub

    Public Function WriteXLSFile(ByVal pDataSet As DataSet) As Boolean
        Try
            'Create a workbook instance
            Dim workbook As Workbook = New Workbook()
            Dim worksheet As Worksheet
            Dim iRow As Integer = 0
            Dim iCol As Integer = 0
            Dim sTemp As String = String.Empty
            Dim dtTemp As DateTime
            Dim dTemp As Double = 0
            Dim iTemp As Integer = 0
            Dim count As Integer = 0
            Dim iTotalRows As Integer = 0
            Dim iSheetCount As Integer = 0

            'Read DataSet
            If Not pDataSet Is Nothing And pDataSet.Tables.Count > 0 Then

                'Traverse DataTable inside the DataSet
                For Each dt As DataTable In pDataSet.Tables

                    'Create a worksheet instance
                    iSheetCount = iSheetCount + 1

                    worksheet = New Worksheet("STOCK INVENTARY REPORT")


                    'Write Table Header
                    iCol = 0
                    For Each dc As DataColumn In dt.Columns
                        worksheet.Cells(0, iCol) = New Cell(dc.ColumnName)
                        iCol = iCol + 1
                    Next

                    'Write Table Body
                    iRow = 1
                    For Each dr As DataRow In dt.Rows
                        iCol = 0
                        For Each dc As DataColumn In dt.Columns
                            sTemp = dr(dc.ColumnName).ToString()
                            Select Case dc.DataType
                                Case GetType(DateTime)
                                    DateTime.TryParse(sTemp, dtTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dtTemp, "MM/DD/YYYY hh:mm:ss")
                                Case GetType(Double)
                                    Double.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.0000")
                                Case GetType(Decimal)
                                    Decimal.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.0000")
                                Case Else
                                    worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                            End Select
                            iCol = iCol + 1
                        Next
                        iRow = iRow + 1
                    Next

                    ''Attach worksheet to workbook
                    workbook.Worksheets.Add(worksheet)
                    iTotalRows = iTotalRows + iRow
                Next

            End If

            'Bug on Excel Library, min file size must be 7 Kb
            'thus we need to add empty row for safety
            If iTotalRows < 100 Then
                worksheet = New Worksheet("Sheet X")
                count = 1
                Do While count < 100
                    worksheet.Cells(count, 0) = New Cell(" ")
                    count = count + 1
                Loop
                workbook.Worksheets.Add(worksheet)
            End If

            'Export the Excel file.
            Response.Clear()
            Response.Buffer = True
            Response.Charset = ""
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=StockInventaryReport.xls")
            Using MyMemoryStream As New MemoryStream()
                workbook.Save(MyMemoryStream)
                MyMemoryStream.WriteTo(Response.OutputStream)
                Response.Flush()
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub showReport(Optional ByVal isPreview As Boolean = False)
        Dim report_table As New DataTable
        Dim rpt_title As String = ""

        Dim sqlString, field, field2 As String
        Dim nDataSource As DataTable

        Dim sumList As New ArrayList

        Dim fileTypeDDL As New DropDownList
        Dim fileType As String = ""

        Try
            If Session("gLang") = "C" Then
                field = "fun_rpt_chi_name"
                field2 = "fun_chi_name"
            Else
                field = "fun_rpt_eng_name"
                field2 = "fun_eng_name"
            End If

            rpt_title = srch_table.Rows(0).Item(field).ToString.Trim

            If rpt_title = "" Then
                rpt_title = srch_table.Rows(0).Item(field2).ToString.Trim

                If Not rpt_title.ToUpper.Contains("REPORT") Then
                    If Session("gLang") = "C" Then
                        rpt_title = rpt_title & " 報表"
                    Else
                        rpt_title = rpt_title & " Report"
                    End If
                End If

            End If

            Dim rptSQL As String = "select *, 'Y' as fldo_field_option from wms_search_col where fun_code = '" & menu_code & "' and srcl_rpt_seq <> 0 order by srcl_rpt_seq"

            If ViewState("ctemp") IsNot Nothing AndAlso ViewState("ctemp") Then
                rptSQL = "select wms_search_col.*, ISNULL(wms_field_option.fldo_field_option, 'Y') as fldo_field_option " &
                        "from wms_search_col left outer join wms_field_option on " &
                        "wms_search_col.cold_tabcol = wms_field_option.FLDO_FIELD_NAME " &
                        "and wms_search_col.FUN_CODE = wms_field_option.FUN_CODE " &
                        "and wms_field_option.imp_code= '" & gU.dbEncode(Session("imp_code")) & "' " &
                        "and wms_field_option.Storer_code= '" & gU.dbEncode(ViewState("sc")) & "' " &
                        "where wms_search_col.fun_code = '" & gU.dbEncode(menu_code) & "' " &
                        "and ISNULL(fldo_field_option,'Y') = 'Y' " &
                        "and srcl_rpt_seq <> 0 " &
                        "order by srcl_rpt_seq "
            End If

            report_table = gDB.getDataTable(rptSQL)

            If isPreview Then
                fileType = "PREVIEW"
            Else
                If srch_table.Rows(0).Item("srch_rpt_filetype_select_yn").ToString.Trim = "Y" Then
                    fileTypeDDL = DirectCast(dtdiv.FindControl("ddlSaveAs"), DropDownList)

                    If fileTypeDDL.SelectedValue.ToString <> "" Then
                        fileType = fileTypeDDL.SelectedItem.Value.ToString.ToUpper
                    Else
                        If Session("gLang") = "C" Then
                            uiFun.displayMsgNew(search_updt_panel, "", "你必須選擇文件類型(下載為)來生成報告!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(search_updt_panel, "", "You must select the Type of File(Download As) to generate Report!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    fileType = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_rpt_default_filetype").ToString.Trim, "PDF")
                End If
            End If

            If report_table Is Nothing Then
                If Not NoAccess Then
                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: no column can be generated. Please check search column.');", True)
                End If

                Exit Sub

            ElseIf report_table.Rows.Count = 0 Then
                If Not NoAccess Then
                    System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: no column can be generated. Please check search column.');", True)
                End If

                Exit Sub
            End If

            Dim srchSQL As String = ""
            Dim whereSQL As String = ""
            Dim criteriaSQL As String = ""
            Dim addSQL As String = ""
            Dim orderbySQL As String = ""
            Dim fieldList As String = ""
            Dim ctrlArray As New Dictionary(Of String, String())

            If srch_table.Rows(0).Item("SRCH_APP_CUSTOMIZED_SQL_YN").ToString.ToUpper.Trim = "Y" Then
                generateSQL(False, "", True, , srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, fieldList, ctrlArray)
                GenericAppMod.ctrlValue = ctrlArray
                sqlString = GenericAppMod.get_ReturnSQL(ModuleType.Search, menu_code, srchSQL, whereSQL, criteriaSQL, addSQL, orderbySQL, fieldList)

            Else
                sqlString = generateSQL(False, "", True)
            End If

            nDataSource = gDB.getDataTable(sqlString)

            Dim r_colName As String = ""
            Dim removeList As New ArrayList
            Dim containList As New ArrayList

            Dim seq As New ArrayList

            'Dim removerptColList As New ArrayList
            'Dim afterSelectFieldStr As String = Session("GENERIC_SESSION_SRCH_SELECT_FIELDS")
            'Dim afterSelectFieldList As String() = Nothing

            'If afterSelectFieldStr IsNot Nothing AndAlso afterSelectFieldStr.Trim <> "" AndAlso afterSelectFieldStr.Contains(",") Then
            '    afterSelectFieldList = Split(afterSelectFieldStr.ToUpper, ",")
            'End If

            'For Each lRows As DataRow In report_table.Select("srcl_rpt_seq <> '0' and isnull(fldo_field_option,'Y') = 'Y'", "srcl_rpt_seq")
            '    Dim FieldName As String = lRows.Item("cold_tabcol").ToString.Trim.ToUpper
            '    Dim colName As String = ""

            '    If FieldName.Contains(".") Then
            '        colName = FieldName.Split(".")(1)
            '    Else
            '        colName = FieldName
            '    End If

            '    If Not nDataSource.Columns.Contains(colName) Then
            '        removerptColList.Add(lRows)
            '    Else
            '        If Not afterSelectFieldStr.Contains("AS " & colName) AndAlso afterSelectFieldList IsNot Nothing AndAlso Not afterSelectFieldList.Contains(FieldName) Then
            '            removerptColList.Add(lRows)
            '        End If
            '    End If
            'Next

            'For Each delRow As DataRow In removerptColList
            '    report_table.Rows.Remove(delRow)
            'Next

            For n As Integer = 0 To report_table.Rows.Count - 1
                seq.Add(n)
            Next
            For i As Integer = 0 To report_table.Rows.Count - 1
                If report_table.Rows(i).Item("cold_tabcol").ToString.Contains(".") Then
                    r_colName = report_table.Rows(i).Item("cold_tabcol").ToString.Trim.Split(".")(1)
                Else
                    r_colName = report_table.Rows(i).Item("cold_tabcol").ToString.Trim
                End If

                If nDataSource.Columns(r_colName) IsNot Nothing Then
                    nDataSource.Columns(r_colName).SetOrdinal(gU.decodeEmptyCInt(seq(i), 0))
                Else
                    If Not NoAccess Then
                        System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('Error: Column " & r_colName & " is not contain in the sql query.');", True)
                    End If

                    Exit Sub
                End If

                containList.Add(r_colName.ToString.ToUpper)
            Next

            For x As Integer = 0 To nDataSource.Columns.Count - 1
                If containList.Contains(nDataSource.Columns(x).ColumnName.ToString.ToUpper) Then
                    Continue For
                Else
                    removeList.Add(nDataSource.Columns(x))
                End If
            Next

            For n As Integer = 0 To removeList.Count - 1
                nDataSource.Columns.Remove(removeList(n))
            Next

            nDataSource.AcceptChanges()
            'generateSumColumn(nDataSource, sumList)

            Dim PageS As New ReportUtils.PageSize
            Dim PageO As New ReportUtils.PageOrentation
            Dim FontSize As String = ""

            FontSize = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_rpt_font_size").ToString.Trim, "8pt")

            Select Case srch_table.Rows(0).Item("srch_rpt_page_size").ToString.Trim
                Case "A4"
                    PageS = ReportUtils.PageSize.A4
                Case "A5"
                    PageS = ReportUtils.PageSize.A5
                Case "A3"
                    PageS = ReportUtils.PageSize.A3
                Case "B4"
                    PageS = ReportUtils.PageSize.B4
                Case Else
                    PageS = ReportUtils.PageSize.A4
            End Select

            Select Case srch_table.Rows(0).Item("srch_rpt_orentation").ToString.Trim
                Case "P"
                    PageO = ReportUtils.PageOrentation.Portrait
                Case "L"
                    PageO = ReportUtils.PageOrentation.Landscape
                Case Else
                    PageO = ReportUtils.PageOrentation.Portrait
            End Select

            RptU.RenderReport(menu_code, nDataSource, report_table, menu_code, fileType,
                                gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_mtable").ToString.Trim, "nulltablename"),
                                Session("gLang"), rpt_title, Me.ClientScript, True, Me, FontSize, PageS, PageO,
                                srch_table.Rows(0).Item("srch_rpt_border_style").ToString.Trim,
                                gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_rpt_show_header_yn").ToString.Trim, "Y"),
                                gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_rpt_show_footer_yn").ToString.Trim, "Y"))

        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(ex.Message) & "');", True)
        End Try
    End Sub

    Private Sub RePopulateCheckBoxes()
        Dim gvrsList As GridView = DirectCast(dtdiv.FindControl(menu_code & "_GridView"), GridView)

        If gvrsList.HeaderRow IsNot Nothing Then
            Dim chkAll As CheckBox = TryCast(gvrsList.HeaderRow.FindControl("chkSelectAll"), CheckBox)

            For Each row As GridViewRow In gvrsList.Rows
                Dim chkBox = TryCast(row.FindControl("btnCB"), CheckBox)
                Dim container As IDataItemContainer = DirectCast(chkBox.NamingContainer, IDataItemContainer)


                If chkAll IsNot Nothing AndAlso chkAll.Checked Then
                    If SelectedCBIndex IsNot Nothing Then
                        If chkBox IsNot Nothing Then
                            Dim cbUnSelected As HiddenField = DirectCast(srchForm.FindControl("cbUnSelected"), HiddenField)

                            Dim keyArray As ArrayList = ViewState("keyArray")
                            Dim valueStr As String = ""
                            If keyArray IsNot Nothing AndAlso keyArray.Count > 0 Then
                                Dim tempConcat As String()

                                For y As Integer = 0 To keyArray.Count - 1
                                    ReDim Preserve tempConcat(y)

                                    tempConcat(y) = gvrsList.DataSource.Rows(gvrsList.PageSize * gvrsList.PageIndex + row.RowIndex).item(keyArray(y)).ToString.Trim
                                Next

                                valueStr = String.Join("#", tempConcat)

                                If valueStr <> "" Then
                                    'Dim cbSelected As HiddenField = DirectCast(srchForm.FindControl("cbSelected"), HiddenField)

                                    If cbUnSelected.Value.Contains(valueStr) Then
                                        chkBox.Checked = False
                                    Else
                                        chkBox.Checked = True
                                    End If
                                End If
                            End If
                        End If

                        'If Not SelectedCBIndex.Exists(Function(i) i = row.RowIndex) Then
                        '    SelectedCBIndex.Add(row.RowIndex)
                        'End If
                    End If
                Else
                    If SelectedCBIndex IsNot Nothing Then
                        If SelectedCBIndex.Exists(Function(i) i = container.DataItemIndex) Then
                            If chkBox IsNot Nothing Then chkBox.Checked = True
                        End If
                    End If

                    If ViewState("dt") IsNot Nothing And ViewState(SELECTED_CB_INDEX) IsNot Nothing Then
                        If ViewState(SELECTED_CB_INDEX).Count = ViewState("dt").Rows.Count Then
                            If chkAll IsNot Nothing Then
                                chkAll.Checked = True
                            End If
                        End If
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub RemoveRowIndex(ByVal index As Integer)
        SelectedCBIndex.Remove(index)
    End Sub

    Private Sub PersistRowIndex(ByVal index As Integer)
        If Not SelectedCBIndex.Exists(Function(i) i = index) Then
            SelectedCBIndex.Add(index)
        End If
    End Sub

    Private Function generateSQL(ByVal isMaster As Boolean, Optional ByVal keyString As String = "", Optional ByVal isReport As Boolean = False,
                                    Optional ByRef form_post As FormPosting = Nothing,
                                    Optional ByRef returnSrchSQL As String = "",
                                    Optional ByRef returnWhereSQL As String = "",
                                    Optional ByRef returnDefCriteriaSQL As String = "",
                                    Optional ByRef returnAddSQL As String = "",
                                    Optional ByRef returnOrderBySQL As String = "",
                                    Optional ByRef returnFieldList As String = "",
                                    Optional ByRef ctlArrayList As Dictionary(Of String, String()) = Nothing,
                                    Optional ByRef isAbort As Boolean = False) As String


        Dim gen_sql As String = ""
        Dim gen_sql_criteria As String = ""
        Dim WhereString As String = ""
        Dim SessionString As String = ""
        Dim sortString As String = ""
        Dim mTableName As String = srch_table.Rows(0).Item("srch_mtable").ToString.Trim

        If Session("usr_id") = "OTSADMIN" Then
            'Throw New Exception("test 1")
        End If

        If srch_table.Rows.Count > 0 Then
            If Not isMaster Then
                If isReport Then
                    generateSQL = srch_table.Rows(0).Item("srch_rpt_sql").ToString.Trim

                    If generateSQL = "" Then generateSQL = srch_table.Rows(0).Item("srch_sql").ToString.Trim
                Else
                    generateSQL = srch_table.Rows(0).Item("srch_sql").ToString.Trim
                End If
            Else
                generateSQL = srch_table.Rows(0).Item("srch_master_sql").ToString.Trim
            End If
        Else
            generateSQL = ""
        End If

        If generateSQL.ToUpper.Contains("WHERE 2=2") Then
            gen_sql_criteria = generateSQL.Substring(generateSQL.ToUpper.IndexOf("WHERE 2=2") + 9, generateSQL.Length - generateSQL.ToUpper.IndexOf("WHERE 2=2") - 9).Trim
            gen_sql = generateSQL.Substring(0, generateSQL.ToUpper.IndexOf("WHERE 2=2") - 1)
        ElseIf generateSQL.ToUpper.Contains("WHERE") Then
            gen_sql_criteria = generateSQL.Substring(generateSQL.ToUpper.IndexOf("WHERE") + 5, generateSQL.Length - generateSQL.ToUpper.IndexOf("WHERE") - 5).Trim
            gen_sql = generateSQL.Substring(0, generateSQL.ToUpper.IndexOf("WHERE") - 1).Trim
        Else
            gen_sql = generateSQL
        End If

        returnFieldList = gen_sql

        If returnFieldList.ToUpper.Contains("SELECT") Then
            returnFieldList = returnFieldList.Substring(returnFieldList.ToUpper.IndexOf("SELECT") + 6, returnFieldList.Length - returnFieldList.ToUpper.IndexOf("SELECT") - 6).Trim
        End If

        If returnFieldList.ToUpper.Contains("DISTINCT") Then
            returnFieldList = returnFieldList.Substring(returnFieldList.ToUpper.IndexOf("DISTINCT") + 8, returnFieldList.Length - returnFieldList.ToUpper.IndexOf("DISTINCT") - 8).Trim
        End If

        If returnFieldList.ToUpper.Contains("FROM") Then
            returnFieldList = returnFieldList.Substring(0, returnFieldList.ToUpper.IndexOf("FROM") - 1).Trim
        End If

        returnFieldList = Replace(returnFieldList, Chr(13), "", , , vbTextCompare)
        returnFieldList = Replace(returnFieldList, Chr(10), "", , , vbTextCompare)
        returnFieldList = returnFieldList.Replace("  ", " ")
        returnFieldList = returnFieldList.Replace(", ", ",")

        If srch_table.Rows(0).Item("SRCH_SELECT_FIELD_YN").ToString.Trim.ToUpper = "Y" Then
            Dim selFieldArray As New ArrayList

            returnFieldList = ""
            selFieldArray = ViewState("SRCH_SELCTED_FIELD_ARRAY")

            If selFieldArray IsNot Nothing AndAlso selFieldArray.Count > 0 Then
                Dim fromSQL As String = ""

                If gen_sql.ToUpper.Contains("FROM") Then
                    fromSQL = gen_sql.Substring(gen_sql.ToUpper.IndexOf("FROM") + 4, gen_sql.Length - gen_sql.ToUpper.IndexOf("FROM") - 4).Trim

                    Dim tmpselSQL As String = ""

                    For x As Integer = 0 To list_table.Rows.Count - 1
                        Dim FieldName As String = list_table.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper

                        If selFieldArray.Contains(FieldName) Then
                            If tmpselSQL <> "" Then tmpselSQL += ","
                            tmpselSQL += FieldName
                        End If
                    Next

                    returnFieldList = tmpselSQL

                    If tmpselSQL <> "" Then
                        gen_sql = "SELECT " & tmpselSQL & " FROM " & fromSQL.ToUpper
                    End If
                End If
            End If
        End If

        returnSrchSQL = gen_sql

        If generateSQL <> "" Then generateSQL = gen_sql & " WHERE "

        If Not isMaster AndAlso keyString <> "" Then
            WhereString = genSearchSQL(keyString, True, form_post, ctlArrayList, isAbort)

            If isAbort Then Exit Function
        Else
            WhereString = genSearchSQL(keyString, False, form_post, ctlArrayList, isAbort)

            If isAbort Then Exit Function
        End If

        SessionString = APP_SESSION_SQL(srch_table.Rows(0).Item("srch_sessioncol").ToString,
                                        srch_table.Rows(0).Item("srch_sessionval").ToString, "AND")

        If gU.decodeNullOrEmpty(WhereString.Trim, "") <> "" Then
            'SV_SRCH_SQL_APP = WhereString
            If WhereString.Contains("WHERE") Then
                WhereString = WhereString.Trim.Substring(WhereString.ToUpper.IndexOf("WHERE") + 5, WhereString.Length - WhereString.ToUpper.IndexOf("WHERE") - 5).Trim
            Else
                WhereString = WhereString.Substring(4, WhereString.Length - 4).Trim
            End If

            If generateSQL <> "" Then generateSQL += WhereString
        Else
            If generateSQL <> "" Then generateSQL += " 1 = 1 "
        End If

        If gU.decodeNullOrEmpty(gen_sql_criteria.Trim, "") <> "" Then
            If generateSQL <> "" Then
                If gen_sql_criteria.Length > 3 AndAlso gen_sql_criteria.Trim.Substring(0, 4).ToUpper = "AND " Then
                    generateSQL += " " & gen_sql_criteria
                Else
                    generateSQL += " AND " & gen_sql_criteria
                End If
            End If

            If WhereString <> "" Then
                If gen_sql_criteria.Length > 3 AndAlso gen_sql_criteria.Trim.Substring(0, 4).ToUpper = "AND " Then
                    returnDefCriteriaSQL += gen_sql_criteria.Trim.Substring(4, gen_sql_criteria.Trim.Length - 4).ToUpper
                Else
                    returnDefCriteriaSQL += gen_sql_criteria.Trim
                End If
            Else
                returnDefCriteriaSQL += gen_sql_criteria.Trim
            End If
        End If

        If gU.decodeNullOrEmpty(SessionString.Trim, "") <> "" Then
            SessionString = SessionString.Substring(4, SessionString.Length - 4)
            If generateSQL <> "" Then generateSQL += " AND " & SessionString

            If WhereString <> "" Then
                WhereString += " AND " & SessionString.Trim
            Else
                WhereString += SessionString
            End If
        End If

        returnWhereSQL = WhereString.Trim

        If Not isMaster Then
            If gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_add_sql").ToString.Trim, "") <> "" Then
                returnAddSQL = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_add_sql").ToString.Trim, "")

                If generateSQL <> "" Then generateSQL += " " & returnAddSQL
            End If
        End If


        If gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_show_sort_yn").ToString.Trim.ToUpper, "Y") = "Y" Then
            Dim orderby As DropDownList = DirectCast(hddiv.FindControl(menu_code & "_orderlist"), DropDownList)

            If orderby.SelectedValue <> "" Then
                Dim value As String = ""
                Dim sortMode As String = ""

                If orderby.SelectedItem.Value.Contains("#") Then
                    value = orderby.SelectedItem.Value.Split("#")(0).Replace("@", ".")

                    Select Case orderby.SelectedItem.Value.Split("#")(1).ToUpper.Trim
                        Case "A"
                            sortMode = ""
                        Case "D"
                            sortMode = "DESC"
                    End Select
                Else
                    value = orderby.SelectedItem.Value.Replace("@", ".")
                End If

                Session("SEARCH_SESSION_PAGE_" & orderby.ID) = value
                If sortString <> "" Then
                    If Not sortString.Contains(value) Then
                        sortString = sortString & ",  " & value & " " & sortMode
                    End If
                Else
                    sortString = value & " " & sortMode
                End If
            Else
                orderby.SelectedIndex = -1
            End If
        End If

        If sortString = "" Then
            Dim sortStr As String = ""

            If isMaster Then
                sortStr = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("SRCH_DEF_MASTER_SORT").ToString.Trim, "")

                If sortStr = "" Then
                    sortStr = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_sort1").ToString.Trim, "")

                    If sortStr <> "" Then
                        sortString = sortStr.Trim
                    End If
                Else
                    sortString = sortStr
                End If
            Else
                sortStr = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("SRCH_DEFAULT_SORT").ToString.Trim, "")

                If sortStr = "" Then
                    sortStr = gU.decodeNullOrEmpty(srch_table.Rows(0).Item("srch_sort1").ToString.Trim, "")

                    If sortStr <> "" Then
                        sortString = sortStr.Trim
                    Else
                        If mTableName <> "" AndAlso Not sortStr.Trim.ToUpper.Contains("ORDER BY") Then
                            If generateSQL.ToUpper.Contains(mTableName.ToUpper & ".SYS_UD") Then
                                sortString = mTableName & ".SYS_UD DESC"
                            End If
                        Else
                            If returnFieldList.ToUpper.Contains("SYS_UD") Then
                                sortString = "SYS_UD DESC"
                            End If
                        End If
                    End If
                Else
                    sortString = sortStr
                End If
            End If
        End If

        If sortString <> "" Then
            If Not sortString.ToUpper.Contains("ORDER BY") Then
                sortString = "ORDER BY " & sortString
            End If

            If generateSQL <> "" Then generateSQL += " " & sortString
        End If

        If sortString.ToUpper <> "" Then
            returnOrderBySQL = sortString
        End If
    End Function

    Private Function genSearchSQL(ByVal keyString As String, Optional ByVal isIgnoePK As Boolean = False,
                                    Optional ByRef form_post As FormPosting = Nothing,
                                    Optional ByRef ctlArrayList As Dictionary(Of String, String()) = Nothing,
                                    Optional ByRef isAbort As Boolean = False) As String

        Session.Remove("SEARCH_SESSION_PAGE_CRITERIA_LIST")

        'Dim symbol As String = srch_table.Rows(0).Item("srch_app_sql").ToString
        Dim table_name As String = srch_table.Rows(0).Item("srch_mtable").ToString
        Dim col_name As String = ""
        Dim dt As New DataTable
        Dim SCString As String = ""

        Try
            If keyString <> "" Then
                SCString = APPSQL(SCString, keyString, "AND")
            End If

            Dim FieldName As String = "'"
            Dim ctrlName As String = ""
            Dim oper As String = ""
            Dim item_srch_sql As String = ""
            Dim isPK As String = ""
            Dim isIgnoingPK As Boolean
            Dim upperEntry As Boolean = False

            Dim symbol As String = "AND"

            For x As Integer = 0 To col_table.Rows.Count - 1
                isIgnoingPK = False
                FieldName = col_table.Rows(x).Item("cold_tabcol").ToString.Trim.ToUpper
                ctrlName = FieldName.Replace(".", "_").Trim
                oper = col_table.Rows(x).Item("srcl_operator").ToString.Trim.ToUpper

                If col_table.Rows(x).Item("srcl_custom_field_yn").ToString.Trim.ToUpper = "Y" Then
                    GenerateObjectSQL(SCString, col_table.Rows(x),
                                            gU.decodeNullOrEmpty(col_table.Rows(x).Item("srcl_edit_style").ToString.Trim, "").ToUpper.Trim,
                                            oper, ctrlName, FieldName, symbol, , form_post, ctlArrayList, isAbort, upperEntry)
                    Continue For
                End If

                If col_table.Rows(x).Item("SRCL_UPPER_ENTRY_YN").ToString.Trim.ToUpper = "Y" Then
                    upperEntry = True
                End If

                If isIgnoePK Then
                    isPK = col_table.Rows(x).Item("SRCL_KEY_YN").ToString.Trim.ToUpper
                    If isPK = "Y" Then
                        isIgnoingPK = True
                    End If
                End If

                If Not isIgnoingPK Then
                    If col_table.Rows(x).Item("SRCL_ISREADONLY_YN").ToString.Trim = "Y" Then
                        If col_table.Rows(x).Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.Trim = "" Then
                            GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString.Trim, oper, ctrlName, FieldName, symbol, , form_post, ctlArrayList, isAbort, upperEntry)

                            If isAbort Then Exit Function
                        Else
                            If col_table.Rows(x).Item("srcl_edit_style").ToString = "S" Then
                                GenerateObjectSQL(SCString, col_table.Rows(x), "H", oper, ctrlName, FieldName, symbol, "S", form_post, ctlArrayList, isAbort, upperEntry)

                                If isAbort Then Exit Function
                            Else
                                'GenerateObjectSQL(SCString, col_table.Rows(x), "L", oper, ctrlName, FieldName, symbol, col_table.Rows(x).Item("srcl_edit_style").ToString.Trim, form_post, ctlArrayList, isAbort)
                                GenerateObjectSQL(SCString, col_table.Rows(x), "Z", oper, ctrlName, FieldName, symbol,
                                                    col_table.Rows(x).Item("srcl_edit_style").ToString.ToUpper.Trim,
                                                    form_post, ctlArrayList, isAbort, upperEntry)

                                If isAbort Then Exit Function
                            End If
                        End If
                    Else
                        GenerateObjectSQL(SCString, col_table.Rows(x), col_table.Rows(x).Item("srcl_edit_style").ToString.Trim, oper, ctrlName, FieldName, symbol, , form_post, ctlArrayList, isAbort, upperEntry)

                        If isAbort Then Exit Function
                    End If

                    If col_table.Rows(x).Item("srcl_filter_req_name").ToString.Trim <> "" Then
                        If Request(col_table.Rows(x).Item("srcl_filter_req_name").ToString.Trim) IsNot Nothing AndAlso
                            Request(col_table.Rows(x).Item("srcl_filter_req_name").ToString.Trim) <> "" Then
                            SCString = APPSQL(SCString, col_table.Rows(x)("cold_tabcol").ToString &
                                                " = '" & gU.dbEncode(Request(col_table.Rows(x).Item("srcl_filter_req_name").ToString.Trim)) & "' ", symbol)
                        End If
                    End If

                End If
            Next

            Dim addDt As New DataTable
            Dim addSQL As String = "select cold_tabcol, srcl_filter_req_name from wms_search_col where fun_code = '" & gU.dbEncode(menu_code) & "' and srcl_filter_req_name is not null"
            addDt = gDB.getDataTable(addSQL)
            For i = 0 To addDt.Rows.Count - 1
                If Request(addDt.Rows(i)("srcl_filter_req_name").ToString) <> "" Then
                    SCString = APPSQL(SCString, addDt.Rows(i)("cold_tabcol").ToString & " = '" & gU.dbEncode(Request(addDt.Rows(i)("srcl_filter_req_name").ToString)) & "' ", symbol)
                End If
            Next

            Dim pfdn As String = ""
            Dim pfdt As String = ""
            Dim pSTORER_CODE As String = ""
            Dim pItem_code As String = ""
            Dim pKey As String = ""
            Dim pILOC_WH As String = ""
            Dim pWH_MAIN_WH As String = ""

            pfdn = ViewState("pfdn")
            pfdt = ViewState("pfdt")

            pSTORER_CODE = ViewState("pSTORER_CODE")
            pILOC_WH = ViewState("pILOC_WH")
            pItem_code = ViewState("pItem_code")
            pKey = ViewState("pPack_key")

            pWH_MAIN_WH = ViewState("pWH_MAIN_WH")

            If pSTORER_CODE <> "" Then
                If table_name <> "" Then
                    If Not SCString.Contains(table_name & ".STORER_CODE") Then _
                            SCString = APPSQL(SCString, " " & table_name & ".STORER_CODE = '" & gU.dbEncode(pSTORER_CODE) & "'", symbol)
                Else
                    If Not SCString.Contains("STORER_CODE") Then _
                            SCString = APPSQL(SCString, " STORER_CODE = '" & gU.dbEncode(pSTORER_CODE) & "'", symbol)
                End If
            End If

            If pILOC_WH <> "" Then
                If menu_code = "LOOKUP_DOWP" Then
                    SCString = APPSQL(SCString, " WMS_DO_PICKLIST_D.PLD_WH = '" & gU.dbEncode(pILOC_WH) & "'", symbol)

                Else
                    If table_name <> "" Then
                        If Not SCString.Contains(table_name & ".ILOC_WH") Then _
                                SCString = APPSQL(SCString, " " & table_name & ".ILOC_WH = '" & gU.dbEncode(pILOC_WH) & "'", symbol)
                    Else
                        If Not SCString.Contains("ILOC_WH") Then _
                                SCString = APPSQL(SCString, " ILOC_WH = '" & gU.dbEncode(pILOC_WH) & "'", symbol)
                    End If
                End If


            End If

            If pWH_MAIN_WH <> "" Then
                Dim tempSQL As String = ""
                Dim tempStr As String = ""
                tempSQL = "Select wh_code from WMS_WAREHOUSE where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and wh_main_wh='" & gU.dbEncode(pWH_MAIN_WH) & "'"
                Dim whDT As DataTable = gDB.getDataTable(tempSQL)

                If whDT.Rows.Count > 0 Then
                    For i = 0 To whDT.Rows.Count - 1
                        tempStr &= "'" & whDT.Rows(i).Item("wh_code").ToString.Trim & "',"
                    Next

                    If tempStr <> "" Then
                        tempStr = Left(tempStr, Len(tempStr) - 1)
                        SCString = APPSQL(SCString, " ILOC_WH in(" & tempStr & ")", symbol)
                    End If

                End If

            End If

            If pItem_code <> "" Then
                If table_name <> "" Then
                    If Not SCString.Contains(table_name & ".ITM_CODE") Then _
                            SCString = APPSQL(SCString, " " & table_name & ".ITM_CODE = '" & gU.dbEncode(pItem_code) & "'", symbol)
                Else
                    If Not SCString.Contains("ITM_CODE") Then _
                            SCString = APPSQL(SCString, " ITM_CODE = '" & gU.dbEncode(pItem_code) & "'", symbol)
                End If
            End If

            If pKey <> "" Then
                If table_name <> "" Then
                    If Not SCString.Contains(table_name & ".PACK_KEY") Then _
                            SCString = APPSQL(SCString, " " & table_name & ".PACK_KEY = '" & gU.dbEncode(pKey) & "'", symbol)
                Else
                    If Not SCString.Contains("PACK_KEY") Then _
                            SCString = APPSQL(SCString, " PACK_KEY = '" & gU.dbEncode(pKey) & "'", symbol)
                End If
            End If


            If pfdt <> "" Then
                If table_name <> "" Then
                    If Not SCString.Contains(table_name & ".LN_DOCTYPE") Then _
                            SCString = APPSQL(SCString, " " & table_name & ".LN_DOCTYPE = '" & gU.dbEncode(pfdt) & "'", symbol)
                Else
                    If Not SCString.Contains("LN_DOCTYPE") Then _
                            SCString = APPSQL(SCString, " LN_DOCTYPE = '" & gU.dbEncode(pfdt) & "'", symbol)
                End If
            End If

            If pfdn <> "" Then
                If table_name <> "" Then
                    If Not SCString.Contains(table_name & ".LN_DOC_NO") Then _
                            SCString = APPSQL(SCString, " " & table_name & ".LN_DOC_NO = '" & gU.dbEncode(pfdn) & "'", symbol)
                Else
                    If Not SCString.Contains("LN_DOC_NO") Then _
                            SCString = APPSQL(SCString, " LN_DOC_NO = '" & gU.dbEncode(pfdn) & "'", symbol)
                End If
            End If

            REM **********************
            'If srch_table.Rows(0).Item("srch_sessioncol").ToString <> "" Then                
            '    Dim sessionList As String()
            '    Dim tempSessionString As String = ""
            '    Dim sessFieldName As String = ""

            '    If srch_table.Rows(0).Item("srch_sessioncol").ToString.Contains(",") Then
            '        sessionList = srch_table.Rows(0).Item("srch_sessioncol").ToString.Split(",")

            '        For j As Integer = 0 To sessionList.Length - 1
            '            sessFieldName = ""

            '            If table_name <> "" Then
            '                sessFieldName = table_name & "." & sessionList(j).ToString.ToUpper.Trim
            '            Else
            '                sessFieldName = sessionList(j).ToString.ToUpper.Trim
            '            End If
            '            tempSessionString = tempSessionString & " AND " & sessFieldName & " = '" & Session(sessionList(j).ToString.Trim) & "'"
            '        Next
            '    Else
            '        If srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim.Contains(".") Then
            '            tempSessionString = " AND " & srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim & " = '" & Session(srch_table.Rows(0).Item("srch_sessioncol").ToString.Trim) & "'"
            '        Else
            '            If table_name <> "" Then
            '                sessFieldName = table_name & "." & srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim
            '            Else
            '                sessFieldName = srch_table.Rows(0).Item("srch_sessioncol").ToString.ToUpper.Trim
            '            End If

            '            tempSessionString = " AND " & sessFieldName & " = '" & Session(srch_table.Rows(0).Item("srch_sessioncol").ToString.Trim) & "'"
            '        End If

            '    End If

            '    SCString = SCString & tempSessionString
            'End If

            REM **********************

            Return SCString

        Catch ex As Exception
            System.Web.UI.ScriptManager.RegisterStartupScript(search_updt_panel, GetType(UpdatePanel), "error", "alert('" & gU.jsString(SCString.ToUpper) & "\n\r\n\r" & gU.jsString(ex.Message) & "');", True)
        Finally
            genSearchSQL = SCString
        End Try
    End Function

    Private Function GenerateCell(ByVal tblRow As DataRow, Optional ByVal rowIndex As Integer = Nothing) As TableCell
        Dim cell As New TableCell()
        Dim customCtrlLink As String() = tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim.Split(",")
        cell.Wrap = False
        Select Case tblRow.Item("srcl_edit_style").ToString.Trim
            Case "T", "CSR"
                If gU.decodeNullOrEmpty(tblRow.Item("srcl_range_yn").ToString.Trim.ToUpper, "N").Trim.ToUpper = "N" Then
                    Dim tb As New TextBox()

                    tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")

                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        tb.CssClass = "REQUIRED"
                    End If

                    If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                        Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                        Dim elseCase As Boolean = True

                        Dim valueQuery = From defKey In defValue
                                         Where defKey.Contains("SESSION")
                                         Select defKey

                        For Each dVal As String In valueQuery
                            tb.Text = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                            elseCase = False
                            Exit For
                        Next

                        If elseCase Then
                            valueQuery = From defKey In defValue
                                         Where defKey.Contains("REQUEST")
                                         Select defKey

                            For Each dVal As String In valueQuery
                                tb.Text = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                                elseCase = False
                                Exit For
                            Next

                        End If

                        If elseCase Then tb.Text = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString

                    End If

                    If remain_status Then
                        tb.Text = Session("SEARCH_SESSION_PAGE_" & tb.ID)
                    End If

                    Dim nUP As New UpdatePanel

                    If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                        nUP.ID = "UP_" & tb.ID
                        AddHandler nUP.Unload, AddressOf UpdatePanel_Unload
                        cell.Controls.Add(nUP)

                        If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                            nUP.ContentTemplateContainer.Controls.Add(tb)
                        End If
                    Else
                        If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                            cell.Controls.Add(tb)
                        End If
                    End If

                    'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                    '    cell.Controls.Add(tb)
                    'End If

                    For j As Integer = 0 To customCtrlLink.Count - 1
                        If customCtrlLink(j) <> "" Then
                            Dim ctrlPlaceHolder As New PlaceHolder
                            ctrlPlaceHolder.ID = "ctrlPlaceHolder" & rowIndex
                            Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                            Dim lit As New Literal

                            userCtrl.ID = "CC_" & rowIndex
                            lit.Text = "&nbsp;"
                            'userCtrl.Attributes.Add("obj_text_name", "dsp_cat_code")
                            userCtrl.Attributes.Add("obj_value_name", tb.ClientID)
                            userCtrl.Attributes.Add("obj_type", TypeName(tb))

                            If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                                userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                            End If

                            ctrlPlaceHolder.Controls.Add(userCtrl)
                            nUP.ContentTemplateContainer.Controls.Add(lit)
                            nUP.ContentTemplateContainer.Controls.Add(ctrlPlaceHolder)
                        End If
                    Next

                    If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                        If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                            nUP.ContentTemplateContainer.Controls.Add(tb)
                        End If
                    Else
                        If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                            cell.Controls.Add(tb)
                        End If
                    End If

                    'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                    '    cell.Controls.Add(tb)
                    'End If
                Else
                    Dim tb1 As New TextBox
                    Dim tb2 As New TextBox
                    Dim lbl1 As New Label
                    Dim lbl2 As New Label

                    tb1.ID = "FR_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb1.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, 150)

                    tb2.ID = "TO_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb2.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, 150)


                    If gU.decodeNullOrEmpty(tblRow.Item("srcl_mandatory_yn").ToString, "").ToUpper = "Y" Then
                        tb1.CssClass = "REQUIRED"
                        tb2.CssClass = "REQUIRED"
                    End If

                    If gU.decodeNullOrEmpty(tblRow.Item("srcl_def_value_sel_seq").ToString, "") <> "" Then
                        Dim defValue As String() = gU.decodeNullOrEmpty(tblRow.Item("srcl_def_value_sel_seq").ToString, "").ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

                        For i As Integer = 0 To defValue.Count - 1
                            If i <= 2 Then
                                If i = 0 Then
                                    If defValue(i).Contains("SESSION") Then
                                        tb1.Text = Session(defValue(i).Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                                    ElseIf defValue(i).Contains("REQUEST") Then
                                        tb1.Text = Request(defValue(i).Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                                    Else
                                        tb1.Text = gU.decodeNullOrEmpty(defValue(i), "")
                                    End If
                                ElseIf i = 1 Then
                                    If defValue(i).Contains("SESSION") Then
                                        tb2.Text = Session(defValue(i).Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                                    ElseIf defValue(i).Contains("REQUEST") Then
                                        tb2.Text = Request(defValue(i).Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                                    Else
                                        tb2.Text = gU.decodeNullOrEmpty(defValue(i), "")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If

                    If remain_status Then
                        tb1.Text = Session("SEARCH_SESSION_PAGE_" & tb1.ID)
                    End If

                    If remain_status Then
                        tb2.Text = Session("SEARCH_SESSION_PAGE_" & tb2.ID)
                    End If

                    If Session("gLang") = "C" Then
                        lbl1.Text = "由 "
                        lbl2.Text = " 至 "
                    Else
                        lbl1.Text = "From "
                        lbl2.Text = " To "
                    End If

                    lbl1.Style.Add("vertical-align", "middle")
                    lbl2.Style.Add("vertical-align", "middle")

                    cell.Controls.Add(lbl1)
                    cell.Controls.Add(tb1)
                    cell.Controls.Add(lbl2)
                    cell.Controls.Add(tb2)
                End If

            Case "D"
                If gU.decodeNullOrEmpty(tblRow.Item("srcl_range_yn").ToString.Trim.ToUpper, "Y").Trim.ToUpper = "N" Then
                    Dim tb As New TextBox

                    tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb.Width = 80
                    tb.MaxLength = 10
                    tb.Style.Add("vertical-align", "middle")
                    tb.Attributes("onkeypress") = "event.returnValue=maskDate(event);"
                    tb.Attributes("onblur") = "if(this.value != '') if(!isDate(this.value, '" & Cache("DDFORMAT2") & "')){this.value=''; alert('Invalid Date!'); this.focus();}"

                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        tb.CssClass = "REQUIRED"
                    End If

                    If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                        If IsNumeric(tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString) Then
                            tb.Text = Format(Now.AddDays(CDbl(tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString)), Cache("DDFORMAT2"))
                        End If
                    End If

                    If remain_status Then
                        tb.Text = Session("SEARCH_SESSION_PAGE_" & tb.ID)
                    End If

                    Dim ajaxCal As New AjaxControlToolkit.CalendarExtender
                    Dim imgBtn As New ImageButton

                    imgBtn.ID = "btn" & tb.ID
                    imgBtn.ImageUrl = "images/calendar1.gif"
                    imgBtn.ImageAlign = ImageAlign.Middle
                    imgBtn.BorderWidth = 0
                    imgBtn.Style.Add("padding-left", "4px")

                    ajaxCal.CssClass = "ajax_calendar"
                    ajaxCal.TargetControlID = tb.ID
                    ajaxCal.PopupButtonID = imgBtn.ID
                    ajaxCal.Format = Cache("DDFORMAT2")

                    cell.Controls.Add(tb)
                    cell.Controls.Add(ajaxCal)
                    cell.Controls.Add(imgBtn)
                Else
                    Dim tb1 As New TextBox
                    Dim tb2 As New TextBox
                    Dim lbl1 As New Label
                    Dim lbl2 As New Label

                    tb1.ID = "FR_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb1.Width = 80
                    tb1.MaxLength = 10
                    tb1.Attributes("onkeypress") = "event.returnValue=maskDate(event);"
                    tb1.Attributes("onblur") = "if(this.value != '') if(!isDate(this.value, '" & Cache("DDFORMAT2") & "')){this.value=''; alert('Invalid Date!'); this.focus();}"

                    If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                        tb1.CssClass = "REQUIRED"
                        tb2.CssClass = "REQUIRED"
                    End If

                    If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                        Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

                        For i As Integer = 0 To defValue.Count - 1
                            If i <= 2 Then
                                If i = 0 Then
                                    tb1.Text = Format(Now.AddDays(CDbl(defValue(i).ToString)), Cache("DDFORMAT2"))
                                ElseIf i = 1 Then
                                    tb2.Text = Format(Now.AddDays(CDbl(defValue(i).ToString)), Cache("DDFORMAT2"))
                                End If
                            End If
                        Next
                    End If

                    If remain_status Then
                        tb1.Text = Session("SEARCH_SESSION_PAGE_" & tb1.ID)
                    End If

                    tb2.ID = "TO_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                    tb2.Width = 80
                    tb2.MaxLength = 10
                    tb2.Attributes("onkeypress") = "event.returnValue=maskDate(event);"
                    tb2.Attributes("onblur") = "if(this.value != '') if(!isDate(this.value, '" & Cache("DDFORMAT2") & "')){this.value=''; alert('Invalid Date!'); this.focus();}"

                    If remain_status Then
                        tb2.Text = Session("SEARCH_SESSION_PAGE_" & tb2.ID)
                    End If

                    If Session("gLang") = "C" Then
                        lbl1.Text = "由 "
                        lbl2.Text = " 至 "
                    Else
                        lbl1.Text = "From "
                        lbl2.Text = " To "
                    End If

                    lbl1.Style.Add("vertical-align", "middle")
                    lbl2.Style.Add("vertical-align", "middle")

                    Dim ajaxCalfr As New AjaxControlToolkit.CalendarExtender
                    Dim ajaxCalto As New AjaxControlToolkit.CalendarExtender
                    Dim imgBtnfr As New ImageButton
                    Dim imgBtnto As New ImageButton

                    imgBtnfr.ImageUrl = "images/calendar1.gif"
                    imgBtnfr.ID = "btn" & tb1.ID
                    imgBtnfr.ImageAlign = ImageAlign.Middle
                    imgBtnfr.BorderWidth = 0
                    imgBtnfr.Style.Add("padding-left", "4px")

                    imgBtnto.ImageUrl = "images/calendar1.gif"
                    imgBtnto.ID = "btn" & tb2.ID
                    imgBtnto.ImageAlign = ImageAlign.Middle
                    imgBtnto.BorderWidth = 0
                    imgBtnto.Style.Add("padding-left", "4px")

                    ajaxCalfr.CssClass = "ajax_calendar"
                    ajaxCalfr.TargetControlID = tb1.ID
                    ajaxCalfr.PopupButtonID = imgBtnfr.ID
                    ajaxCalfr.Format = Cache("DDFORMAT2")

                    ajaxCalto.CssClass = "ajax_calendar"
                    ajaxCalto.TargetControlID = tb2.ID
                    ajaxCalto.PopupButtonID = imgBtnto.ID
                    ajaxCalto.Format = Cache("DDFORMAT2")

                    cell.Controls.Add(lbl1)
                    cell.Controls.Add(tb1)
                    cell.Controls.Add(ajaxCalfr)
                    cell.Controls.Add(imgBtnfr)
                    cell.Controls.Add(lbl2)
                    cell.Controls.Add(tb2)
                    cell.Controls.Add(ajaxCalto)
                    cell.Controls.Add(imgBtnto)
                End If

            Case "C"
                Dim cbl As New CheckBoxList
                cbl.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                cbl.CssClass = "NEWTEXT"

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    cbl.CssClass = "REQUIRED"
                End If

                If tblRow.Item("srcl_checkbox_rowcnt").ToString <> "" Then _
                    cbl.RepeatColumns = CInt(tblRow.Item("srcl_checkbox_rowcnt").ToString)

                cbl.RepeatDirection = RepeatDirection.Horizontal

                Dim li As ListItem

                Dim cbString As String = ""

                Dim cbTbl As New DataTable
                Dim gotSeq As Boolean = False

                If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                    cbString = "select colc_code as code, colc_eng_value as name, colc_chi_value as cname, colc_display_seq from wms_col_code where colc_tabcol = '" &
                                            tblRow.Item("cold_tabcol").ToString.Trim & "' " &
                                            "order by colc_display_seq"
                    gotSeq = True
                Else
                    cbString = tblRow.Item("srcl_sql").ToString.Trim
                End If

                cbTbl = gDB.getDataTable(cbString)

                Dim tempCheckedList As String() = Nothing

                If remain_status Then
                    If Session("SEARCH_SESSION_PAGE_" & cbl.ID) IsNot Nothing Then
                        tempCheckedList = Session("SEARCH_SESSION_PAGE_" & cbl.ID).ToString.Split(",")
                    End If
                End If

                Dim defArrayList As New ArrayList
                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)

                    Dim valueQuery = From defKey In defValue
                                     Select defKey

                    For Each dVal As String In valueQuery
                        defArrayList.Add(dVal)
                    Next
                End If

                For x As Integer = 0 To cbTbl.Rows.Count - 1
                    li = New ListItem

                    If Session("gLang") = "C" AndAlso cbTbl.Columns.Contains("cname") Then
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

                    If remain_status Then
                        If Not tempCheckedList Is Nothing Then
                            If tempCheckedList.Contains(cbTbl.Rows(x).Item("code").ToString) Then
                                li.Selected = True
                            End If
                        End If
                    End If

                    cbl.Items.Insert(x, li)
                Next

                cell.Controls.Add(cbl)

            Case "S"
                Dim ddl As New DropDownList
                ddl.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    ddl.CssClass = "REQUIRED"
                End If

                If tblRow.Item("srcl_sql").ToString.Trim = "" Then
                    uiFun.load_dropdownBy_ColCode(ddl, tblRow.Item("cold_tabcol").ToString.Trim, Session("gLang"))
                Else
                    Dim app_symbol As String = tblRow.Item("srcl_app_sql").ToString
                    Dim ddlSQL As String = tblRow.Item("srcl_sql").ToString.Trim

                    ddlSQL += " " & APP_SESSION_SQL(tblRow.Item("srcl_sessioncol").ToString, tblRow.Item("srcl_sessionval").ToString, app_symbol)
                    ddlSQL += " " & APP_REQUEST_SQL(tblRow.Item("srcl_requestcol").ToString, tblRow.Item("srcl_requestval").ToString, app_symbol)

                    If gU.decodeNullOrEmpty(tblRow.Item("srcl_add_sql").ToString, "").Trim <> "" Then
                        ddlSQL += " " & tblRow.Item("srcl_add_sql").ToString.Trim.ToUpper
                    End If

                    If Not ddlSQL.ToUpper.Contains("ORDER BY") Then
                        ddlSQL += " ORDER BY 2"
                    End If

                    uiFun.load_dropdown(ddl, ddlSQL.Trim, , , , Session("gSelectLabel"))
                End If

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue
                                     Where defKey.Contains("SESSION")
                                     Select defKey

                    For Each dVal As String In valueQuery
                        ddl.SelectedValue = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then
                        valueQuery = From defKey In defValue
                                     Where defKey.Contains("REQUEST")
                                     Select defKey

                        For Each dVal As String In valueQuery
                            ddl.SelectedValue = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                            elseCase = False
                            Exit For
                        Next
                    End If

                    If elseCase Then ddl.SelectedValue = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString
                End If

                If remain_status Then
                    ddl.SelectedValue = Session("SEARCH_SESSION_PAGE_" & ddl.ID)
                End If

                Dim nUP As New UpdatePanel

                If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                    nUP.ID = "UP_" & ddl.ID
                    AddHandler nUP.Unload, AddressOf UpdatePanel_Unload
                    cell.Controls.Add(nUP)
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(ddl)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        cell.Controls.Add(ddl)
                    End If
                End If

                'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                '    cell.Controls.Add(ddl)
                'End If

                For j As Integer = 0 To customCtrlLink.Count - 1
                    If customCtrlLink(j) <> "" Then
                        Dim ctrlPlaceHolder As New PlaceHolder
                        ctrlPlaceHolder.ID = "ctrlPlaceHolder" & rowIndex
                        Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                        Dim lit As New Literal

                        userCtrl.ID = "CC_" & rowIndex
                        lit.Text = "&nbsp;"
                        'userCtrl.Attributes.Add("obj_text_name", "dsp_cat_code")
                        userCtrl.Attributes.Add("obj_value_name", ddl.ClientID)
                        userCtrl.Attributes.Add("obj_type", TypeName(ddl))

                        If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                            userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                        End If

                        ctrlPlaceHolder.Controls.Add(userCtrl)
                        nUP.ContentTemplateContainer.Controls.Add(lit)
                        nUP.ContentTemplateContainer.Controls.Add(ctrlPlaceHolder)
                    End If
                Next

                If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(ddl)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        cell.Controls.Add(ddl)
                    End If
                End If

                'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                '    cell.Controls.Add(ddl)
                'End If

            Case "L"
                Dim objValueHD As New HiddenField
                Dim lblHiddenField As New HiddenField
                'Dim lbl As New Label
                Dim lbl As New TextBox

                'lbl.ReadOnly = True
                lbl.BorderStyle = BorderStyle.None
                lbl.BackColor = Drawing.Color.Transparent
                lbl.Style.Add("vertical-align", "Middle")

                objValueHD.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                lblHiddenField.ID = "dsphd_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                lbl.ID = "dsp_" & tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                lbl.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")
                'lbl.Height = 20

                lbl.Attributes.Add("onkeydown", "return false;")
                lbl.Attributes.Add("onchange", "this.value = document.getElementById('" & lblHiddenField.ClientID & "').value;")

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    'cell.CssClass = "REQUIRED"
                    'lbl.Text = "&nbsp;"
                    lbl.BackColor = Drawing.ColorTranslator.FromHtml("#FFFA9C")
                End If

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue
                                     Where defKey.Contains("SESSION")
                                     Select defKey

                    For Each dVal As String In valueQuery
                        objValueHD.Value = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then
                        valueQuery = From defKey In defValue
                                     Where defKey.Contains("REQUEST")
                                     Select defKey

                        For Each dVal As String In valueQuery
                            objValueHD.Value = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                            elseCase = False
                            Exit For
                        Next
                    End If

                    If elseCase Then
                        objValueHD.Value = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString
                    End If
                End If

                If remain_status Then
                    objValueHD.Value = Session("SEARCH_SESSION_PAGE_" & objValueHD.ID)
                End If

                If lblHiddenField.Value.Trim <> "" Then
                    lbl.Text = lblHiddenField.Value.Trim
                ElseIf ViewState(lbl.ID & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(lbl.ID & "_VIEWSTATE").ToString <> "" Then
                    lblHiddenField.Value = ViewState(lbl.ID & "_VIEWSTATE")
                    lbl.Text = lblHiddenField.Value.Trim
                End If

                If remain_status Then
                    lblHiddenField.Value = Session("SEARCH_SESSION_PAGE_" & lbl.ID)
                    lbl.Text = Session("SEARCH_SESSION_PAGE_" & lbl.ID)
                End If

                Dim nUP As New UpdatePanel

                If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                    nUP.ID = "UP_" & lbl.ID
                    AddHandler nUP.Unload, AddressOf UpdatePanel_Unload
                    cell.Controls.Add(nUP)
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(objValueHD)
                        nUP.ContentTemplateContainer.Controls.Add(lbl)
                        nUP.ContentTemplateContainer.Controls.Add(lblHiddenField)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        cell.Controls.Add(objValueHD)
                        cell.Controls.Add(lbl)
                        cell.Controls.Add(lblHiddenField)
                    End If
                End If

                For j As Integer = 0 To customCtrlLink.Count - 1
                    If customCtrlLink(j) <> "" Then
                        Dim ctrlPlaceHolder As New PlaceHolder
                        ctrlPlaceHolder.ID = "ctrlPlaceHolder" & rowIndex
                        Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                        Dim lit As New Literal

                        userCtrl.ID = "CC_" & rowIndex
                        lit.Text = "&nbsp;"
                        userCtrl.Attributes.Add("obj_text_name_hidden", lblHiddenField.ClientID)
                        userCtrl.Attributes.Add("obj_text_name", lbl.ClientID)

                        userCtrl.Attributes.Add("obj_value_name", objValueHD.ClientID)
                        userCtrl.Attributes.Add("obj_type", TypeName(objValueHD))

                        If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                            userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                        End If

                        ctrlPlaceHolder.Controls.Add(userCtrl)
                        'cell.Controls.Add(lit)
                        'cell.Controls.Add(ctrlPlaceHolder)

                        nUP.ContentTemplateContainer.Controls.Add(lit)
                        nUP.ContentTemplateContainer.Controls.Add(ctrlPlaceHolder)
                    End If
                Next

                If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(objValueHD)
                        nUP.ContentTemplateContainer.Controls.Add(lbl)
                        nUP.ContentTemplateContainer.Controls.Add(lblHiddenField)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        cell.Controls.Add(objValueHD)
                        cell.Controls.Add(lbl)
                        cell.Controls.Add(lblHiddenField)
                    End If
                End If

                'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                '    cell.Controls.Add(objValueHD)
                '    cell.Controls.Add(lbl)
                '    cell.Controls.Add(lblHiddenField)

                'End If

                'lbl.ReadOnly = True

            Case "H"
                Dim hf As New HiddenField

                hf.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue
                                     Where defKey.Contains("SESSION")
                                     Select defKey

                    For Each dVal As String In valueQuery
                        hf.Value = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then
                        valueQuery = From defKey In defValue
                                     Where defKey.Contains("REQUEST")
                                     Select defKey

                        For Each dVal As String In valueQuery
                            hf.Value = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                            elseCase = False
                            Exit For
                        Next

                    End If

                    If elseCase Then hf.Value = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString
                End If

                Dim nUP As New UpdatePanel

                If remain_status Then
                    hf.Value = Session("SEARCH_SESSION_PAGE_" & hf.ID)
                End If

                If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then

                    If tblRow.Item("srcl_app_custom_ctrl_links").ToString.Trim <> "" Then
                        nUP.ID = "UP_" & hf.ID

                        AddHandler nUP.Unload, AddressOf UpdatePanel_Unload
                        cell.Controls.Add(nUP)
                        nUP.ContentTemplateContainer.Controls.Add(hf)
                    Else
                        cell.Controls.Add(hf)
                    End If


                    For j As Integer = 0 To customCtrlLink.Count - 1
                        If customCtrlLink(j) <> "" Then
                            Dim ctrlPlaceHolder As New PlaceHolder
                            ctrlPlaceHolder.ID = "ctrlPlaceHolder" & rowIndex

                            Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                            userCtrl.ID = "CC_" & rowIndex
                            userCtrl.Attributes.Add("obj_value_name", hf.ClientID)
                            userCtrl.Attributes.Add("obj_type", TypeName(hf))

                            If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                                userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                            End If

                            ctrlPlaceHolder.Controls.Add(userCtrl)
                            nUP.ContentTemplateContainer.Controls.Add(ctrlPlaceHolder)
                        End If
                    Next
                Else
                    srchForm.Controls.Add(hf)
                End If

            Case "U"
                Dim tb As New TextBox

                tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                tb.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")
                tb.Style.Add("vertical-align", "middle")
                cell.Controls.Add(tb)

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    tb.CssClass = "REQUIRED"
                End If

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    Dim defValue As String() = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString.ToUpper.Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                    Dim elseCase As Boolean = True

                    Dim valueQuery = From defKey In defValue
                                     Where defKey.Contains("SESSION")
                                     Select defKey

                    For Each dVal As String In valueQuery
                        tb.Text = Session(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("SESSION", ""))
                        elseCase = False
                        Exit For
                    Next

                    If elseCase Then
                        valueQuery = From defKey In defValue
                                     Where defKey.Contains("REQUEST")
                                     Select defKey

                        For Each dVal As String In valueQuery
                            tb.Text = Request(dVal.Replace("""", "").Replace("(", "").Replace(")", "").Replace("REQUEST", ""))
                            elseCase = False
                            Exit For
                        Next
                    End If

                    If elseCase Then tb.Text = tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString

                End If

                If remain_status Then
                    tb.Text = Session("SEARCH_SESSION_PAGE_" & tb.ID)
                End If

                If tblRow.Item("srcl_lookup_fun_code").ToString.Trim <> "" Then
                    Dim imgBtn As New ImageButton
                    Dim newFormLit As New Literal
                    Dim w_height As Integer = 600
                    Dim w_width As Integer = 500
                    Dim w_left As Integer = 5
                    Dim w_top As Integer = 15

                    If tblRow.Item("srcl_lookup_win_width").ToString.Trim <> "" Then w_height = gU.decodeEmptyCInt(tblRow.Item("srcl_lookup_win_width").ToString.Trim, 0)
                    If tblRow.Item("srcl_lookup_win_height").ToString.Trim <> "" Then w_width = gU.decodeEmptyCInt(tblRow.Item("srcl_lookup_win_height").ToString.Trim, 0)
                    If tblRow.Item("srcl_lookup_win_left").ToString.Trim <> "" Then w_left = gU.decodeEmptyCInt(tblRow.Item("srcl_lookup_win_left").ToString.Trim, 0)
                    If tblRow.Item("srcl_lookup_win_top").ToString.Trim <> "" Then w_top = gU.decodeEmptyCInt(tblRow.Item("srcl_lookup_win_top").ToString.Trim, 0)

                    imgBtn.Attributes.Add("onmouseout", "MM_swapImgRestore()")
                    imgBtn.ID = "Image_" & tb.ID & "_LookUp"
                    imgBtn.Attributes.Add("onmousedown", "MM_swapImage('" & imgBtn.ClientID & "','','images/btn_search_over.gif',1)")
                    imgBtn.Style.Add("cursor", "hand")
                    imgBtn.ImageUrl = "images/btn_search.gif"
                    imgBtn.ImageAlign = ImageAlign.Middle
                    imgBtn.OnClientClick = "javascript:" & tb.ID & "LOOKUP();event.returnValue=false;"

                    cell.Controls.Add(imgBtn)

                    newFormLit.Text = "<form name=""hidden_" & tb.ID & """ method=""POST""></form>"
                    Me.Page.Controls.Add(newFormLit)

                    Dim jsStr As String = ""

                    jsStr += "<script language=""Javascript"">" & vbNewLine
                    jsStr += "function " & tb.ID & "LOOKUP() {" & vbNewLine
                    jsStr += "  removeAllElementFromForm(document.hidden_" & tb.ID & ");" & vbNewLine
                    jsStr += "  window.open("""", """ & tblRow.Item("srcl_lookup_fun_code").ToString.Trim & ""","
                    jsStr += """titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,"
                    jsStr += "width=" & w_width & ","
                    jsStr += "height=" & w_height & ","
                    jsStr += "left=" & w_left & ","
                    jsStr += "top=" & w_top
                    jsStr += """);" & vbNewLine
                    jsStr += "  setInterfaceDataToForm(document.hidden_" & tb.ID & ", ""menu_code"", """ & tblRow.Item("srcl_lookup_fun_code").ToString.Trim & """);" & vbNewLine
                    jsStr += "  setInterfaceDataToForm(document.hidden_" & tb.ID & ", ""pForm"", """ & Me.Form.ClientID & """);" & vbNewLine
                    jsStr += "  setInterfaceDataToForm(document.hidden_" & tb.ID & ", ""pItemList"", """ & tb.ClientID & """);" & vbNewLine
                    jsStr += "  document.hidden_" & tb.ID & ".action = ""cms_search.aspx"";" & vbNewLine
                    jsStr += "  document.hidden_" & tb.ID & ".target = """ & tblRow.Item("srcl_lookup_fun_code").ToString.Trim & """;" & vbNewLine
                    jsStr += "  document.hidden_" & tb.ID & ".submit();" & vbNewLine
                    jsStr += "}" & vbNewLine
                    jsStr += "</script>"

                    ClientScript.RegisterStartupScript(Me.GetType(), tb.ID & "_LOOKUP", jsStr)
                End If
            Case "B"
                Return Nothing
            Case Else
                Dim tb As New TextBox

                tb.ID = tblRow.Item("cold_tabcol").ToString.Trim.Replace(".", "_")
                tb.Width = gU.decodeEmptyCInt(tblRow.Item("srcl_col_width").ToString, "120")

                If tblRow.Item("srcl_mandatory_yn").ToString.ToUpper = "Y" Then
                    tb.CssClass = "REQUIRED"
                End If

                If remain_status Then
                    tb.Text = Session("SEARCH_SESSION_PAGE_" & tb.ID)
                End If

                Dim nUP As New UpdatePanel

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    nUP.ID = "UP_" & tb.ID

                    AddHandler nUP.Unload, AddressOf UpdatePanel_Unload
                    cell.Controls.Add(nUP)

                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(tb)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                        cell.Controls.Add(tb)
                    End If
                End If

                'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper = "Y" Then
                '    cell.Controls.Add(tb)
                'End If

                For j As Integer = 0 To customCtrlLink.Count - 1
                    If customCtrlLink(j) <> "" Then
                        Dim ctrlPlaceHolder As New PlaceHolder
                        ctrlPlaceHolder.ID = "ctrlPlaceHolder" & rowIndex
                        Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                        Dim lit As New Literal

                        userCtrl.ID = "CC_" & rowIndex
                        lit.Text = "&nbsp;"
                        'userCtrl.Attributes.Add("obj_text_name", "dsp_cat_code")
                        userCtrl.Attributes.Add("obj_value_name", tb.ClientID)
                        userCtrl.Attributes.Add("obj_type", TypeName(tb))

                        If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                            userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                        End If

                        ctrlPlaceHolder.Controls.Add(userCtrl)
                        nUP.ContentTemplateContainer.Controls.Add(lit)
                        nUP.ContentTemplateContainer.Controls.Add(ctrlPlaceHolder)
                    End If
                Next

                If tblRow.Item("SRCL_DEF_VALUE_SEL_SEQ").ToString <> "" Then
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        nUP.ContentTemplateContainer.Controls.Add(tb)
                    End If
                Else
                    If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                        cell.Controls.Add(tb)
                    End If
                End If

                'If tblRow.Item("srcl_rend_custom_ctrl_atlast").ToString.Trim.ToUpper <> "Y" Then
                '    cell.Controls.Add(tb)
                'End If
        End Select

        cell.VerticalAlign = VerticalAlign.Middle
        GenerateCell = cell
    End Function

    Private Function APPSQL(ByVal SQL As String, ByVal Clause As String, ByVal SYMBOL As String) As String

        APPSQL = ""

        Select Case SYMBOL.ToUpper
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

    Private Function APP_SESSION_SQL(session_col As String, session_val As String, Optional symbol As String = "AND") As String
        Dim sessionSQL As String = ""

        If session_col <> "" AndAlso session_val <> "" Then
            Dim sessFieldName As String = ""

            If session_col.Contains(",") AndAlso session_val.Contains(",") Then
                Dim sessionColList As String()
                Dim sessionValList As String()
                Dim replaceDict As New Dictionary(Of String, String)
                Dim sessFieldValue As String = ""

                sessionColList = session_col.Split(",")
                sessionValList = session_val.Split(",")

                For j As Integer = 0 To sessionColList.Length - 1
                    If sessionColList(j).ToString.Contains("#") Then
                        sessFieldName = sessionColList(j).ToString.ToUpper.Trim.Substring(0, sessionColList(j).ToString.ToUpper.Trim.IndexOf("#"))

                        If Not replaceDict.Keys.Contains(sessionColList(j).ToString.ToUpper.Trim.Substring(sessionColList(j).ToString.ToUpper.Trim.IndexOf("#"), _
                                                                                                            sessionColList(j).ToString.Trim.Length - sessionColList(j).ToString.Trim.IndexOf("#"))) Then
                            replaceDict.Add(sessionColList(j).ToString.ToUpper.Trim.Substring(sessionColList(j).ToString.ToUpper.Trim.IndexOf("#"), _
                                                                                                sessionColList(j).ToString.Trim.Length - sessionColList(j).ToString.Trim.IndexOf("#")), sessFieldName)
                        End If
                    Else
                        sessFieldName = sessionColList(j).ToString.ToUpper.Trim
                    End If

                    sessFieldValue = sessionValList(j).ToString.ToUpper.Trim

                    If replaceDict IsNot Nothing AndAlso replaceDict.Keys.Count > 0 Then
                        If Session(sessFieldValue) IsNot Nothing AndAlso Session(sessFieldValue) <> "" Then
                            Dim tmpStr As String = Session(sessFieldValue)
                            For k As Integer = 0 To replaceDict.Keys.Count - 1
                                tmpStr = tmpStr.Replace(replaceDict.Keys(k), replaceDict.Item(replaceDict.Keys(k)))
                            Next

                            If sessFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                sessionSQL = APPSQL(sessionSQL, sessFieldName & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                            Else
                                sessionSQL = APPSQL(sessionSQL, "ISNULL(" & sessFieldName & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                            End If
                        End If
                    Else
                        If sessFieldName <> "" AndAlso Session(sessFieldValue) IsNot Nothing Then
                            If Session(sessFieldValue) IsNot Nothing AndAlso Session(sessFieldValue) <> "" Then

                                If sessFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                    sessionSQL = APPSQL(sessionSQL, sessFieldName & "='" & gU.dbEncode(Session(sessFieldValue)) & "'", symbol)
                                Else
                                    sessionSQL = APPSQL(sessionSQL, "ISNULL(" & sessFieldName & ",'') = '" & gU.dbEncode(Session(sessFieldValue)) & "'", symbol)
                                End If
                            End If
                        ElseIf sessFieldName = "" AndAlso sessFieldValue <> "" Then
                            If Session(sessFieldValue) IsNot Nothing AndAlso Session(sessFieldValue) <> "" Then
                                sessionSQL = APPSQL(sessionSQL, Session(sessFieldValue), symbol)
                            End If
                        End If
                    End If
                Next
            Else
                If session_col.Contains("#") Then
                    Dim key As String

                    sessFieldName = session_col.Trim.Substring(0, session_col.Trim.IndexOf("#"))
                    key = session_col.Trim.Substring(session_col.Trim.IndexOf("#"), session_col.Trim.Length - session_col.Trim.IndexOf("#"))
                    If Session(session_val) IsNot Nothing Then
                        sessionSQL = APPSQL(sessionSQL, Session(session_val).ToString.Replace(key, sessFieldName), symbol)
                    End If
                Else
                    If session_col <> "" AndAlso Session(session_val) IsNot Nothing Then
                        If Session(session_val) IsNot Nothing AndAlso Session(session_val) <> "" Then
                            If session_col.Trim.ToUpper <> "COMP_CODE" Then
                                sessionSQL = APPSQL(sessionSQL, session_col & "='" & gU.dbEncode(Session(session_val)) & "'", symbol)
                            Else
                                sessionSQL = APPSQL(sessionSQL, "ISNULL(" & session_col & ",'') = '" & gU.dbEncode(Session(session_val)) & "'", symbol)
                            End If
                        End If
                    ElseIf session_col = "" AndAlso session_val <> "" Then
                        If Session(session_val) IsNot Nothing AndAlso Session(session_val) <> "" Then
                            sessionSQL = APPSQL(sessionSQL, Session(session_val), symbol)
                        End If
                    End If
                End If
            End If

        ElseIf session_col = "" AndAlso session_val <> "" Then
            If Session(session_val) IsNot Nothing AndAlso Session(session_val) <> "" Then
                sessionSQL = APPSQL(sessionSQL, Session(session_val), symbol)
            End If
        End If

        Return sessionSQL
    End Function

    Private Function APP_REQUEST_SQL(request_col As String, request_val As String, Optional symbol As String = "AND") As String
        Dim requestSQL As String = ""

        If request_col <> "" AndAlso request_val <> "" Then
            Dim reqFieldName As String = ""

            If request_col.Contains(",") AndAlso request_val.Contains(",") Then
                Dim requestColList As String()
                Dim requestValList As String()
                Dim replaceDict As New Dictionary(Of String, String)
                Dim reqFieldValue As String = ""

                requestColList = request_col.Split(",")
                requestValList = request_val.Split(",")

                For j As Integer = 0 To requestColList.Length - 1
                    If requestColList(j).ToString.Contains("#") Then
                        reqFieldName = requestColList(j).ToString.ToUpper.Trim.Substring(0, requestColList(j).ToString.ToUpper.Trim.IndexOf("#"))

                        If Not replaceDict.Keys.Contains(requestColList(j).ToString.ToUpper.Trim.Substring(requestColList(j).ToString.ToUpper.Trim.IndexOf("#"), _
                                                                                                            requestColList(j).ToString.Trim.Length - requestColList(j).ToString.Trim.IndexOf("#"))) Then
                            replaceDict.Add(requestColList(j).ToString.ToUpper.Trim.Substring(requestColList(j).ToString.ToUpper.Trim.IndexOf("#"), _
                                                                                                requestColList(j).ToString.Trim.Length - requestColList(j).ToString.Trim.IndexOf("#")), reqFieldName)
                        End If
                    Else
                        reqFieldName = requestColList(j).ToString.ToUpper.Trim
                    End If

                    reqFieldValue = requestValList(j).ToString.ToUpper.Trim

                    If replaceDict IsNot Nothing AndAlso replaceDict.Keys.Count > 0 Then
                        If Request(reqFieldValue) IsNot Nothing AndAlso Request(reqFieldValue) <> "" Then
                            Dim tmpStr As String = Request(reqFieldValue)

                            ViewState(reqFieldValue & "_VIEWSTATE") = tmpStr

                            For k As Integer = 0 To replaceDict.Keys.Count - 1
                                tmpStr = tmpStr.Replace(replaceDict.Keys(k), replaceDict.Item(replaceDict.Keys(k)))
                            Next

                            If reqFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                requestSQL = APPSQL(requestSQL, reqFieldName & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                            Else
                                requestSQL = APPSQL(requestSQL, "ISNULL(" & reqFieldName & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                            End If

                        ElseIf ViewState(reqFieldValue & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(reqFieldValue & "_VIEWSTATE") <> "" Then
                            Dim tmpStr As String = ViewState(reqFieldValue & "_VIEWSTATE")

                            For k As Integer = 0 To replaceDict.Keys.Count - 1
                                tmpStr = tmpStr.Replace(replaceDict.Keys(k), replaceDict.Item(replaceDict.Keys(k)))
                            Next

                            If reqFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                requestSQL = APPSQL(requestSQL, reqFieldName & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                            Else
                                requestSQL = APPSQL(requestSQL, "ISNULL(" & reqFieldName & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                            End If

                        End If
                    Else
                        If reqFieldName <> "" AndAlso Request(reqFieldValue) IsNot Nothing Then
                            If Request(reqFieldValue) IsNot Nothing AndAlso Request(reqFieldValue) <> "" Then
                                Dim tmpStr As String = Request(reqFieldValue)

                                ViewState(reqFieldValue & "_VIEWSTATE") = tmpStr

                                If reqFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                    requestSQL = APPSQL(requestSQL, reqFieldName & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                                Else
                                    requestSQL = APPSQL(requestSQL, "ISNULL(" & reqFieldName & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                                End If
                            End If
                        ElseIf reqFieldName <> "" AndAlso ViewState(reqFieldValue & "_VIEWSTATE") IsNot Nothing Then
                            If ViewState(reqFieldValue & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(reqFieldValue & "_VIEWSTATE") <> "" Then
                                Dim tmpStr As String = ViewState(reqFieldValue & "_VIEWSTATE")

                                If reqFieldValue.Trim.ToUpper <> "COMP_CODE" Then
                                    requestSQL = APPSQL(requestSQL, reqFieldName & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                                Else
                                    requestSQL = APPSQL(requestSQL, "ISNULL(" & reqFieldName & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                                End If
                            End If
                        ElseIf reqFieldName = "" AndAlso reqFieldValue <> "" Then
                            If Request(reqFieldValue) IsNot Nothing AndAlso Request(reqFieldValue) <> "" Then
                                Dim tmpStr As String = Request(reqFieldValue)

                                ViewState(reqFieldValue & "_VIEWSTATE") = tmpStr

                                requestSQL = APPSQL(requestSQL, tmpStr, symbol)
                            ElseIf ViewState(reqFieldValue & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(reqFieldValue & "_VIEWSTATE") <> "" Then
                                Dim tmpStr As String = ViewState(reqFieldValue & "_VIEWSTATE")

                                requestSQL = APPSQL(requestSQL, tmpStr, symbol)
                            End If
                        End If
                    End If
                Next
            Else
                If request_col.Contains("#") Then
                    Dim key As String

                    reqFieldName = request_col.Trim.Substring(0, request_col.Trim.IndexOf("#"))
                    key = request_col.Trim.Substring(request_col.Trim.IndexOf("#"), request_col.Trim.Length - request_col.Trim.IndexOf("#"))

                    If Request(request_val) IsNot Nothing Then
                        Dim tmpStr As String = Request(request_val)

                        ViewState(request_val & "_VIEWSTATE") = tmpStr

                        requestSQL = APPSQL(requestSQL, tmpStr.ToString.Replace(key, reqFieldName), symbol)

                    ElseIf ViewState(request_val & "_VIEWSTATE") IsNot Nothing Then
                        Dim tmpStr As String = ViewState(request_val & "_VIEWSTATE")

                        requestSQL = APPSQL(requestSQL, tmpStr.ToString.Replace(key, reqFieldName), symbol)
                    End If
                Else
                    If request_col <> "" AndAlso Request(request_val) IsNot Nothing Then
                        If Request(request_val) IsNot Nothing AndAlso Request(request_val) <> "" Then
                            Dim tmpStr As String = Request(request_val)

                            ViewState(request_val & "_VIEWSTATE") = tmpStr

                            If request_col.Trim.ToUpper <> "COMP_CODE" Then
                                requestSQL = APPSQL(requestSQL, request_col & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                            Else
                                requestSQL = APPSQL(requestSQL, "ISNULL(" & request_col & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                            End If
                        End If
                    ElseIf request_col <> "" AndAlso ViewState(request_val & "_VIEWSTATE") IsNot Nothing Then
                        If ViewState(request_val & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(request_val & "_VIEWSTATE") <> "" Then
                            Dim tmpStr As String = ViewState(request_val & "_VIEWSTATE")

                            If request_col.Trim.ToUpper <> "COMP_CODE" Then
                                requestSQL = APPSQL(requestSQL, request_col & "='" & gU.dbEncode(tmpStr) & "'", symbol)
                            Else
                                requestSQL = APPSQL(requestSQL, "ISNULL(" & request_col & ",'') = '" & gU.dbEncode(tmpStr) & "'", symbol)
                            End If
                        End If
                    ElseIf request_col = "" AndAlso request_val <> "" Then
                        If Request(request_val) IsNot Nothing AndAlso Request(request_val) <> "" Then
                            Dim tmpStr As String = Request(request_val)

                            ViewState(request_val & "_VIEWSTATE") = tmpStr

                            requestSQL = APPSQL(requestSQL, tmpStr, symbol)

                        ElseIf ViewState(request_val & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(request_val & "_VIEWSTATE") <> "" Then
                            Dim tmpStr As String = ViewState(request_val & "_VIEWSTATE")

                            requestSQL = APPSQL(requestSQL, tmpStr, symbol)
                        End If
                    End If
                End If
            End If

        ElseIf request_col = "" AndAlso request_val <> "" Then
            If Request(request_val) IsNot Nothing AndAlso Request(request_val) <> "" Then
                Dim tmpStr As String = Request(request_val)

                ViewState(request_val & "_VIEWSTATE") = tmpStr

                requestSQL = APPSQL(requestSQL, tmpStr, symbol)

            ElseIf ViewState(request_val & "_VIEWSTATE") IsNot Nothing AndAlso ViewState(request_val & "_VIEWSTATE") <> "" Then
                Dim tmpStr As String = ViewState(request_val & "_VIEWSTATE")

                requestSQL = APPSQL(requestSQL, tmpStr, symbol)
            End If
        End If

        Return requestSQL
    End Function

    Private Function GetSortDirection(ByVal column As String) As String

        ' By default, set the sort direction to ascending.
        Dim sortDirection = "ASC"

        ' Retrieve the last column that was sorted.
        Dim sortExpression = TryCast(ViewState("SortExpression"), String)

        If sortExpression IsNot Nothing Then
            ' Check if the same column is being sorted.
            ' Otherwise, the default value can be returned.
            If sortExpression = column Then
                Dim lastDirection = TryCast(ViewState("SortDirection"), String)
                If lastDirection IsNot Nothing _
                    AndAlso lastDirection = "ASC" Then

                    sortDirection = "DESC"

                End If
            End If
        End If

        ' Save new values in ViewState.
        ViewState("SortDirection") = sortDirection
        ViewState("SortExpression") = column

        Return sortDirection

    End Function

    Private Sub buildAttribute(ByRef obj As Object, ByVal RowDataItem As Object, ByVal keySet As ICollection, Optional ByVal usekeyCode As Boolean = False)

        For a As Integer = 0 To keySet.Count - 1
            Dim att_key As String = keySet(a)
            Dim att_val As String = ""

            att_val = obj.Attributes(att_key)

            If att_val <> "" AndAlso att_val.Contains("{") AndAlso att_val.Contains("}") Then
                Dim keyArray As New ArrayList

                If usekeyCode Then
                    keyArray = ViewState("keyCodeArray")
                Else
                    keyArray = ViewState("keyArray")
                End If

                If keyArray IsNot Nothing AndAlso keyArray.Count > 0 Then
                    Dim args As String()

                    For y As Integer = 0 To keyArray.Count - 1
                        ReDim Preserve args(y)

                        args(y) = HttpUtility.UrlEncode(gU.dbEncode(DataBinder.Eval(RowDataItem, keyArray(y)).ToString.Trim))
                    Next

                    obj.Attributes(att_key) = String.Format(att_val, args)
                End If
            End If
        Next
    End Sub

    Protected Sub UpdatePanel_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        RegisterUpdatePanel(sender)
    End Sub

    Private Sub RegisterUpdatePanel(ByVal Panel As UpdatePanel)
        For Each methodInfo As MethodInfo In GetType(ScriptManager).GetMethods(BindingFlags.NonPublic Or BindingFlags.Instance)

            If methodInfo.Name.Equals("System.Web.UI.IScriptManagerInternal.RegisterUpdatePanel") Then
                methodInfo.Invoke(ScriptManager.GetCurrent(Page), New Object() {Panel})

            End If
        Next
    End Sub

    Private Function validateObj(ByVal validateValue As String, ByVal type As String, Optional datatype As String = "") As Boolean
        Select Case type.ToUpper.Trim
            Case "D"
                If validateValue <> "" AndAlso Not gU.isValidDate(validateValue) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(search_updt_panel, "", "Invalid date! " & validateValue, Session("gLang"))
                    Else
                        uiFun.displayMsgNew(search_updt_panel, "", "無效的日期! " & validateValue, Session("gLang"))
                    End If

                    Return False
                End If
            Case Else
                If datatype = "NUMBER" Then
                    If validateValue <> "" AndAlso Not IsNumeric(validateValue) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(search_updt_panel, "", "Invalid Number! " & validateValue, Session("gLang"))
                        Else
                            uiFun.displayMsgNew(search_updt_panel, "", "無效的Number! " & validateValue, Session("gLang"))
                        End If

                        Return False
                    End If
                End If
        End Select

        Return True
    End Function

    <System.Web.Services.WebMethod()> _
    Public Shared Sub removeTemplateSession()
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Class GridViewLabelTemplate
        Inherits System.Web.UI.Page
        Implements ITemplate

        Private ItemType As ListItemType
        Private columnName As String
        Private ControlID As String
        Private ControlName As String
        Private elementType As String
        Private HDText As String
        Private hyperURL As String
        Private js_key As String
        Private js_value As String
        Private ControlText As String
        Private customCtrl_Link As String
        Private customCtrl_UPanel As UpdatePanel

        Public Sub New(ByVal type As ListItemType, Optional ByVal CtrlID As String = "", Optional ByVal CtrlName As String = "", _
                                                    Optional ByVal eleType As String = "", Optional ByVal colname As String = "", _
                                                    Optional ByVal HeaderText As String = "", Optional ByVal hyperlinkURL As String = "", _
                                                    Optional ByVal js_attrib_key As String = "", Optional ByVal js_attrib_value As String = "", _
                                                    Optional ByVal CtrlText As String = "", Optional ByVal cc_link As String = "", Optional ByRef cc_updatePanel As UpdatePanel = Nothing)
            ItemType = type
            columnName = colname
            ControlID = CtrlID
            ControlName = CtrlName
            elementType = eleType
            HDText = HeaderText
            hyperURL = hyperlinkURL
            js_key = js_attrib_key
            js_value = js_attrib_value
            ControlText = CtrlText
            customCtrl_Link = cc_link
            customCtrl_UPanel = cc_updatePanel
        End Sub

        Private Sub InstantiateIn(ByVal container As System.Web.UI.Control) Implements ITemplate.InstantiateIn
            Select Case ItemType
                Case ListItemType.Header
                    Select Case elementType
                        Case "MCHECKBOX"
                            Dim chkAll As New CheckBox
                            chkAll.ID = ControlID

                            If js_key <> "" Then chkAll.Attributes.Add(js_key, "javascript:" & js_value)

                            container.Controls.Add(chkAll)

                        Case "LINKBUTTON"
                            Dim Hd As New Literal
                            Hd.Text = HDText
                            container.Controls.Add(Hd)
                        Case Else
                            Exit Select
                    End Select

                Case ListItemType.Item
                    Select Case elementType
                        Case "MCHECKBOX"
                            Dim cb As New CheckBox
                            cb.ID = ControlID
                            If js_key <> "" Then cb.Attributes.Add(js_key, "javascript:" & js_value)
                            container.Controls.Add(cb)

                            Dim hf As New HiddenField
                            hf.ID = "VALUE"
                            container.Controls.Add(hf)

                            Dim hf1 As New HiddenField
                            hf1.ID = "VALUE1"
                            container.Controls.Add(hf1)

                            Dim hf2 As New HiddenField
                            hf2.ID = "VALUE2"
                            container.Controls.Add(hf2)

                            Dim cbhf As New HiddenField
                            cbhf.ID = ControlID & "_HD"
                            container.Controls.Add(cbhf)

                        Case "HTMLINPUTBUTTON"
                            Dim btn As New HtmlInputButton
                            btn.ID = ControlID
                            btn.Attributes.Add("class", "all_button")
                            btn.Style.Add("font-Size", "11px")
                            btn.Style.Add("height", "22px")
                            btn.Style.Add("width", "50px")
                            btn.Name = ControlName
                            btn.Value = ControlText
                            If js_key <> "" Then btn.Attributes.Add(js_key, "javascript:" & js_value)
                            container.Controls.Add(btn)

                        Case "CHECKBOX"
                            Dim cb As New CheckBox
                            cb.ID = ControlID
                            If js_key <> "" Then cb.Attributes.Add(js_key, "javascript:" & js_value)
                            container.Controls.Add(cb)

                        Case "HF"
                            Dim hf As New HiddenField
                            hf.ID = ControlID
                            container.Controls.Add(hf)

                        Case "LABEL"
                            Dim lbl As New Label
                            lbl.ID = ControlID
                            lbl.Text = ControlText
                            lbl.CssClass = "GV"
                            lbl.Style.Add("font-Size", "11px")
                            container.Controls.Add(lbl)

                        Case "LINKBUTTON"
                            Dim LinkB As New LinkButton
                            LinkB.ID = ControlID
                            LinkB.CssClass = "GV"
                            LinkB.Text = ControlText
                            LinkB.Style.Add("font-Size", "11px")
                            LinkB.CommandName = "Update"
                            'AddHandler LinkB.Click, AddressOf mt.LinkMaster
                            If js_key <> "" Then LinkB.Attributes.Add(js_key, "javascript:" & js_value)
                            container.Controls.Add(LinkB)

                        Case "HTMLANCHOR"
                            Dim hLink As New HtmlAnchor
                            hLink.ID = ControlID
                            hLink.Name = ControlName
                            hLink.InnerText = ControlText

                            hLink.Style.Add("font-Size", "11px")
                            'hLink.CssClass = "GV"
                            If js_key <> "" Then hLink.Attributes.Add(js_key, "javascript:" & js_value)
                            hLink.HRef = hyperURL
                            container.Controls.Add(hLink)
                        Case "CUSTOMCONTROL"
                            Dim customCtrlLink As String() = customCtrl_Link.Split(",")

                            For j As Integer = 0 To customCtrlLink.Count - 1
                                If customCtrlLink(j) <> "" Then
                                    Dim ctrlPlaceHolder As New PlaceHolder
                                    ctrlPlaceHolder.ID = "ctrlPlaceHolder_" & ControlID
                                    Dim userCtrl As UserControl = Page.LoadControl(customCtrlLink(j))
                                    Dim lit As New Literal

                                    userCtrl.ID = ControlID
                                    lit.Text = "&nbsp;"
                                    'userCtrl.Attributes.Add("obj_text_name", "dsp_cat_code")
                                    'userCtrl.Attributes.Add("obj_value_name", tb.ClientID)
                                    'userCtrl.Attributes.Add("obj_type", TypeName(tb))

                                    'If tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim <> "" Then
                                    '    userCtrl.Attributes.Add("passing_value", tblRow.Item("srcl_value_pass_to_custom_ctrl").ToString.Trim)
                                    'End If

                                    ctrlPlaceHolder.Controls.Add(userCtrl)
                                    container.Controls.Add(ctrlPlaceHolder)
                                End If
                            Next

                        Case "DELETEBUTTON"
                            Dim btn As New Button
                            btn.CssClass = "all_button"
                            btn.Font.Bold = False
                            REM============================================================
                            btn.CommandName = "Delete"   REM# GridView CommandName Eevent: 
                            REM#                              Delete = gridview.RowDeleting
                            REM#                              Edit = gridview.RowEditing
                            REM#                              Update = gridview.RowUpdating
                            REM============================================================
                            btn.Text = ControlText
                            btn.ID = ControlID
                            btn.Style.Add("font-Size", "11px")
                            btn.Style.Add("height", "22px")
                            btn.Style.Add("width", "50px")
                            'AddHandler btn.Click, AddressOf btnDeleteClick
                            container.Controls.Add(btn)
                        Case Else
                            Exit Select
                    End Select
                Case ListItemType.Footer
                    Exit Select
                Case Else
                    Exit Select
            End Select
        End Sub

        Private Sub btnDeleteClick(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class

    Private Class GridViewPagerTemplate
        Inherits System.Web.UI.Page
        Implements ITemplate

        Private ItemType As ListItemType
        Private pagerDDL As New DropDownList
        Private previousBtn As ImageButton
        Private nextBtn As ImageButton
        Private firstBtn As ImageButton
        Private lastBtn As ImageButton

        Public Sub New(ByVal type As ListItemType, ByRef pager_dropdown As DropDownList, _
                                                    ByRef previous_btn As ImageButton, _
                                                    ByRef next_btn As ImageButton, _
                                                    ByRef first_btn As ImageButton, _
                                                    ByRef last_btn As ImageButton)
            ItemType = type
            pagerDDL = pager_dropdown
            previousBtn = previous_btn
            nextBtn = next_btn
            firstBtn = first_btn
            lastBtn = last_btn
        End Sub

        Private Sub InstantiateIn(ByVal container As System.Web.UI.Control) Implements ITemplate.InstantiateIn
            Select Case ItemType
                Case ListItemType.Pager
                    Dim pager_label As New Label

                    pager_label.ID = "lblPager"
                    pager_label.Text = "Page : "
                    pager_label.Font.Bold = True
                    pager_label.Style.Add("vertical-align", "middle")

                    container.Controls.Add(firstBtn)
                    container.Controls.Add(previousBtn)
                    container.Controls.Add(pager_label)
                    container.Controls.Add(pagerDDL)
                    container.Controls.Add(nextBtn)
                    container.Controls.Add(lastBtn)
                Case Else
                    Exit Select
            End Select
        End Sub
    End Class

    Private Class FormPosting
        Private Inputs As New System.Collections.Specialized.NameValueCollection()
        Public Url As String = ""
        Public Method As String = "post"
        Public FormName As String = "submitForm"
        Public FormTarget As String = "_self"

        Public Sub Add(ByVal name As String, ByVal value As String)
            Inputs.Add(name, value)
        End Sub

        Public Sub Post()
            System.Web.HttpContext.Current.Response.Clear()
            System.Web.HttpContext.Current.Response.Write("<html><head></head>")

            System.Web.HttpContext.Current.Response.Write(String.Format("<body onload=""document.{0}.submit();"">", FormName))
            System.Web.HttpContext.Current.Response.Write(String.Format("<form name=""{0}"" method=""{1}"" action=""{2}"" target=""{3}"" style=""width:0;height:0;"">", FormName, Method, Url, FormTarget))

            For i As Integer = 0 To Inputs.Keys.Count - 1
                System.Web.HttpContext.Current.Response.Write(String.Format("<input name=""{0}"" type=""hidden"" value=""{1}"">", Inputs.Keys(i), Inputs(Inputs.Keys(i))))
            Next

            System.Web.HttpContext.Current.Response.Write("</form>")
            System.Web.HttpContext.Current.Response.Write("</body></html>")
            System.Web.HttpContext.Current.Response.End()
        End Sub
    End Class
End Class


