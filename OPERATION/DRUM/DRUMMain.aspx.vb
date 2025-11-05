Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_STA_DRUMMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans

    Private DDFORMAT As String = "DD/MM/YYYY"
    Private exceptionEditList As List(Of String)
    Private moduleAction As String = ""
    Private dt As New DataTable

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

            Dim tImage As Image = Image_Loc_LookUp_DRUM_LOC
            tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
            tImage.Attributes.Add("onclick", "LocLookUp(1,'" & HttpUtility.HtmlEncode(dsp_DRUM_LOC.ClientID) & "', '" & HttpUtility.HtmlEncode(DRUM_LOC.ClientID) & "')")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(DRUM_WH_CODE, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))

        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If

            If DRUM_DATE.Text = "" Then
                DRUM_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Drum / Trolley Maintenance"
            lbl_ImageHd.Text = "Movement Details"
            lbl_DRUM_ID.Text = "Drum ID:"
            lbl_DRUM_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_DRUM_TYPE.Text = "Type:"
            lbl_DRUM_DATE.Text = "Date:"
            lbl_DRUM_WH_CODE.Text = "Base Warehouse:"
            lbl_DRUM_REMARKS.Text = "Remarks:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Inactive"
            newrow.Text = "Add"

            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"



        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨量調整維護"
            lbl_ImageHd.Text = "貨量調整詳情"

            lbl_DRUM_ID.Text = "調整編號:"
            lbl_DRUM_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_DRUM_TYPE.Text = "類型:"

            lbl_DRUM_DATE.Text = "日期:"

            lbl_DRUM_WH_CODE.Text = "倉庫:"
            lbl_DRUM_REMARKS.Text = "備註"

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
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"


        End If
        REM **********************

        REM **********************
        REM Additional CSS

        DRUM_WH_CODE.CssClass = "REQUIRED"
        DRUM_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'AD_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("DRUM_ID") = ""
            ViewState("STORER_CODE") = ""
            Call BindGV()

            DRUM_IS_EMPTY.Text = checkIsemptyDrum()

        Else
            dt = ViewState("dt")
            reloadGridView()
        End If

        'If moduleAction = "SELECTIM" Then
        '    addItemtoSTA()
        'End If

        'cm = New CommonMenu("SADJ", lheader.text, AD_CODE.Text)
        'cm.parentDir = "../../"
        'cm.haveCheckList = "Y"
        'cm.haveAttachments = "Y"
        'cm.haveNotes = "Y"
        'cm.haveTasks = "Y"
        'cm.haveEmail = "Y"
        'cm.haveHistory = "Y"

        'cm.genCM(cmBar)
        setPageCtrlAccess()

        selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value, document.myform." & DRUM_WH_CODE.ClientID & ".value);return false;")
        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OP_DRUM','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("DRUM_ID") & "','N');")

        If DRUM_STATUS.Text = "INACTIVE" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
    End Sub
    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("btnReOpen")

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
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Code.", "物料編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物料名稱")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Loc.", "位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Orig. Qty", "原來數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Revised Qty", "修訂數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Variance Qty", "差異數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "Expiry Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "ManuFactory Date", "ManuFactory Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "", "")
        '        REM **********************

        '        oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("mFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("dsp_DRUD_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "FULL_DRUD_DATE").ToString.Trim
                CType(e.Row.FindControl("DRUD_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_DATE").ToString.Trim
                CType(e.Row.FindControl("DRUD_DATE_HR"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_DATE_HR").ToString.Trim
                CType(e.Row.FindControl("DRUD_DATE_MIN"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_DATE_MIN").ToString.Trim

                CType(e.Row.FindControl("DRUD_MV_TYPE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "DRUD_MV_TYPE").ToString.Trim
                CType(e.Row.FindControl("DRUD_MOVEMENT"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_MOVEMENT").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("DRUD_DOC_TYPE"), DropDownList), "Select COLC_CODE, COLC_ENG_VALUE from wms_col_code where COLC_TABCOL='WMS_DRUM_DTL.DRUD_DOC_TYPE' order by COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_CODE", , , DataBinder.Eval(e.Row.DataItem, "DRUD_DOC_TYPE").ToString.Trim)

                'CType(e.Row.FindControl("DRUD_DOC_TYPE"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "DRUD_DOC_TYPE").ToString.Trim
                CType(e.Row.FindControl("DRUD_DOC_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_DOC_NO").ToString.Trim
                CType(e.Row.FindControl("DRUD_BY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_BY").ToString.Trim
                CType(e.Row.FindControl("DRUD_WH"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_WH").ToString.Trim

                CType(e.Row.FindControl("dsp_DRUD_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_LOC").ToString.Trim
                CType(e.Row.FindControl("DRUD_LOC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "DRUD_LOC").ToString.Trim

                Dim tImage As Image = CType(e.Row.FindControl("Image_DRUD_LOC"), Image)
                tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                tImage.Attributes.Add("onclick", "LocLookUp('','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_DRUD_LOC"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("DRUD_LOC"), HiddenField).ClientID) & "')")

                CType(e.Row.FindControl("DRUD_REMARKS"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUD_REMARKS").ToString.Trim

                'Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)
                'If Session("gLang") = "E" Then
                '    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                '    nButton.Text = "Delete"
                'ElseIf Session("gLang") = "C" Then
                '    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                '    nButton.Text = "删除"
                'End If

                If DataBinder.Eval(e.Row.DataItem, "DRUD_MV_TYPE").ToString.Trim <> "MANUAL" Then
                    CType(e.Row.FindControl("dsp_DRUD_LOC"), Label).Visible = True
                    CType(e.Row.FindControl("DRUD_DATE"), TextBox).Visible = False
                    CType(e.Row.FindControl("DRUD_DATE_HR"), TextBox).Visible = False
                    CType(e.Row.FindControl("DRUD_DATE_MIN"), TextBox).Visible = False
                    CType(e.Row.FindControl("btnDATE_ID1"), ImageButton).Visible = False
                    CType(e.Row.FindControl("DRUD_DOC_TYPE"), DropDownList).Enabled = False
                    CType(e.Row.FindControl("DRUD_DOC_NO"), TextBox).Enabled = False
                    CType(e.Row.FindControl("DRUD_BY"), TextBox).Enabled = False
                    CType(e.Row.FindControl("DRUD_REMARKS"), TextBox).Enabled = False
                Else
                    CType(e.Row.FindControl("dsp_DRUD_DATE"), Label).Visible = False
                    CType(e.Row.FindControl("DRUD_DATE"), TextBox).Visible = True
                    CType(e.Row.FindControl("DRUD_DATE_HR"), TextBox).Visible = True
                    CType(e.Row.FindControl("DRUD_DATE_MIN"), TextBox).Visible = True
                    CType(e.Row.FindControl("btnDATE_ID1"), ImageButton).Visible = True
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
        'Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        'dt.Rows(e.RowIndex).Item("mFlag") = "D"
        'dt.AcceptChanges()
    End Sub

    Private Sub reloadGridView()
        If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("dsp_DRUD_LOC"), Label).Text = CType(GridView1.Rows(i).FindControl("DRUD_LOC"), HiddenField).Value
            Next
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        If DRUM_ID.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_DRUM_ID.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_DRUM_ID.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        Else
            If Session("pagemode") = "N" Then
                Dim existsCount As Integer = 0

                selectSql = "select count(*) from WMS_DRUM " & _
                             "where DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' " & _
                             "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                             "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"

                existsCount = gU.decodeEmptyCInt(DB.getValueFromSQL(selectSql), 0)

                If existsCount > 0 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Duplicate Drum ID has found! Please enter another ID", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "Duplicate Drum ID has found! Please enter another ID", Session("gLang"))
                    End If
                    Return False
                End If
            End If
        End If




        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If DRUM_WH_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_DRUM_WH_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_DRUM_WH_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If DRUM_LOC.Value = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Master Location cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_DRUM_LOC.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If DRUM_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_DRUM_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_DRUM_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(DRUM_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_DRUM_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_DRUM_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim dtlType As String = CType(GridView1.Rows(i).FindControl("DRUD_MV_TYPE"), HiddenField).Value

                If dtlType = "MANUAL" Then
                    If CType(GridView1.Rows(i).FindControl("DRUD_DATE"), TextBox).Text.Trim = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please input Movement Date!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Please input Movement Date!", Session("gLang"))
                        End If
                        Return False
                    Else
                        If Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("DRUD_DATE"), TextBox).Text.Trim) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Invalid Movement Date!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "Invalid Movement Date!", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If


                    If CType(GridView1.Rows(i).FindControl("DRUD_DOC_TYPE"), DropDownList).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select Document type!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Please select Document type!", Session("gLang"))
                        End If
                        Return False
                    End If

                    'If CType(GridView1.Rows(i).FindControl("DRUD_DOC_NO"), TextBox).Text.Trim = "" Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Please enter Document No.!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "Please enter Document No.!", Session("gLang"))
                    '    End If
                    '    Return False
                    'End If

                    If CType(GridView1.Rows(i).FindControl("DRUD_LOC"), HiddenField).Value = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select movement Location!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Please select movement Location!", Session("gLang"))
                        End If
                        Return False
                    End If
                End If
            Next
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
        Dim nextSEQ As Integer

        Dim successFlag As Boolean = False

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try

                Dim dateArr As String()
                Dim dateStr As String = ""
                Dim monStr, dayStr, yearStr As String

                Dim isResrv As String = ""
                Dim isInsp As String = ""

                If DRUM_IS_RESERVED.Checked Then isResrv = "Y" Else isResrv = "N"
                If DRUM_IS_INSP.Checked Then isInsp = "Y" Else isInsp = "N"


                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DRUM_ID.Text.Trim
                    'nextNo = AD_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from WMS_DRUM " & _
                                "where DRUM_ID = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into WMS_DRUM (DRUM_ID, IMP_CODE, STORER_CODE,  DRUM_TYPE, DRUM_DATE, DRUM_WH_CODE, DRUM_LOC, DRUM_IN_WH, DRUM_CATEGORY, DRUM_DESC, " & _
                        " DRUM_REMARKS, DRUM_SERIAL_NO, STATUS, DRUM_IS_RESERVED, DRUM_IS_INSP, SYS_LUB, SYS_LUD,SYS_CB, SYS_CD )" & _
                        " values ( " & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(DRUM_TYPE.SelectedValue)) & ", " & gU.convdbDate(gU.dbEncode(DRUM_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DRUM_WH_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(DRUM_LOC.Value)) & "," & gU.convdbNVCData(gU.dbEncode(DRUM_IN_WH.SelectedValue)) & ",NULL," & _
                        gU.convdbNVCData(gU.dbEncode(DRUM_DESC.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DRUM_REMARKS.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(DRUM_SERIAL_NO.Text.Trim)) & ",'ACTIVE'," & _
                        gU.convdbNVCData(isResrv) & "," & gU.convdbNVCData(isInsp) & "," & _
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        REM **********************s

                        ViewState("DRUM_ID") = nextNo

                        dt = ViewState("dt")

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "DRUD_SEQ")

                            For Each rows As DataRow In dt.Rows

                                REM **********************
                                REM Modify Here
                                If rows.Item("DRUD_MV_TYPE").ToString.Trim = "MANUAL" Then
                                    dateStr = ""
                                    monStr = ""
                                    dayStr = ""
                                    yearStr = ""
                                    Select Case rows.Item("mFlag")
                                        Case "N"
                                            nextSEQ = gU.decodeEmptyCInt(DB.getValueFromSQL("Select max(isnull(DRUD_SEQ,0)) + 1 from WMS_DRUM_DTL Where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                      "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                      "and DRUM_ID = '" & gU.dbEncode(nextNo) & "'", gConn, transaction), 1)

                                            If rows.Item("DRUD_DATE").ToString.Trim <> "" Then
                                                dateArr = Split(rows.Item("DRUD_DATE").ToString.Trim, "/")

                                                If dateArr.Length = 3 Then
                                                    dateStr = " SMALLDATETIMEFROMPARTS ('" & dateArr(2) & "','" & dateArr(1) & "','" & dateArr(0) & "','" & gU.decodeNullOrEmpty(rows.Item("DRUD_DATE_HR").ToString.Trim, "0") & "','" & gU.decodeNullOrEmpty(rows.Item("DRUD_DATE_MIN").ToString.Trim, "0") & "') "
                                                End If

                                            End If

                                            itemSQL = " INSERT INTO WMS_DRUM_DTL (DRUM_ID, IMP_CODE, STORER_CODE, DRUD_SEQ, DRUD_DATE, DRUD_MOVEMENT, DRUD_DOC_TYPE, DRUD_DOC_NO, DRUD_BY, DRUD_LOC, " & _
                                                      " DRUD_REMARKS, DRUD_MV_TYPE, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                                      " VALUES (" & gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(nextSEQ)) & ", " & gU.decodeNullOrEmpty(dateStr, "NULL") & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_MOVEMENT").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_DOC_TYPE").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_DOC_NO").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_BY").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_LOC").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_REMARKS").ToString.Trim, ""))) & ", " & _
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_MV_TYPE").ToString.Trim, ""))) & ", " & _
                                                      "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    End Select
                                    REM **********************

                                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                                End If
                            Next
                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate Dum ID has been found. Please enter another Drum ID!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨量調整資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_drum set " & _
                                    "DRUM_date = " & gU.convdbDate(gU.dbEncode(DRUM_DATE.Text.Trim)) & ", " & _
                                    "DRUM_type = " & gU.convdbNVCData(gU.dbEncode(DRUM_TYPE.SelectedValue)) & ", " & _
                                    "DRUM_SERIAL_NO = " & gU.convdbNVCData(gU.dbEncode(DRUM_SERIAL_NO.Text.Trim)) & ", " & _
                                    "DRUM_WH_CODE = " & gU.convdbNVCData(gU.dbEncode(DRUM_WH_CODE.SelectedValue)) & ", " & _
                                    "DRUM_REMARKS = " & gU.convdbNVCData(gU.dbEncode(DRUM_REMARKS.Text.Trim)) & ", " & _
                                    "DRUM_LOC = " & gU.convdbNVCData(gU.dbEncode(DRUM_LOC.Value)) & ", " & _
                                    "DRUM_IN_WH = " & gU.convdbNVCData(gU.dbEncode(DRUM_IN_WH.SelectedValue)) & ", " & _
                                    "DRUM_DESC = " & gU.convdbNVCData(gU.dbEncode(DRUM_DESC.Text.Trim)) & ", " & _
                                    "DRUM_IS_RESERVED=" & gU.convdbNVCData(isResrv) & "," & _
                                    "DRUM_IS_INSP=" & gU.convdbNVCData(isInsp) & "," & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text) & "' "

                    REM **********************

                    dt = ViewState("dt")

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "DRUD_SEQ")

                        For Each rows As DataRow In dt.Rows
                            dateStr = ""
                            monStr = ""
                            dayStr = ""
                            yearStr = ""

                            itemSQL = ""

                            If rows.Item("DRUD_DATE").ToString.Trim <> "" Then
                                dateArr = Split(rows.Item("DRUD_DATE").ToString.Trim, "/")
                                If dateArr.Length = 3 Then
                                    dateStr = " SMALLDATETIMEFROMPARTS ('" & dateArr(2) & "','" & dateArr(1) & "','" & dateArr(0) & "','" & gU.decodeNullOrEmpty(rows.Item("DRUD_DATE_HR").ToString.Trim, "0") & "','" & gU.decodeNullOrEmpty(rows.Item("DRUD_DATE_MIN").ToString.Trim, "0") & "') "
                                End If
                            End If


                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"

                                    nextSEQ = gU.decodeEmptyCInt(DB.getValueFromSQL("Select max(isnull(convert(int,DRUD_SEQ),0)) + 1 from WMS_DRUM_DTL Where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "'", gConn, transaction), 1)

                                    itemSQL = " INSERT INTO WMS_DRUM_DTL (DRUM_ID, IMP_CODE, STORER_CODE, DRUD_SEQ, DRUD_DATE, DRUD_MOVEMENT, DRUD_DOC_TYPE, DRUD_DOC_NO, DRUD_BY, DRUD_LOC, " & _
                                              " DRUD_REMARKS, DRUD_MV_TYPE, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                              " VALUES        (" & gU.convdbNVCData(gU.dbEncode(DRUM_ID.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(nextSEQ)) & ", " & gU.decodeNullOrEmpty(dateStr, "NULL") & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_MOVEMENT").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_DOC_TYPE").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_DOC_NO").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_BY").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_LOC").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_REMARKS").ToString.Trim, ""))) & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUD_MV_TYPE").ToString.Trim, ""))) & ", " & _
                                              "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    'itemSQL = "delete from wms_stock_adjust_d " & _
                                    '        "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    '        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    '        "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' " & _
                                    '        "and DRUD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("DRUD_SEQ").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update WMS_DRUM_DTL set " & _
                                                "DRUD_DATE= " & gU.decodeNullOrEmpty(dateStr, "NULL") & ", " & _
                                                "DRUD_MOVEMENT = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_MOVEMENT").ToString.Trim, ""))) & ", " & _
                                                "DRUD_DOC_TYPE = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_DOC_TYPE").ToString.Trim, ""))) & ", " & _
                                                "DRUD_DOC_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_DOC_NO").ToString.Trim, ""))) & ", " & _
                                                "DRUD_BY = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_BY").ToString.Trim, ""))) & ", " & _
                                                "DRUD_LOC = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_LOC").ToString.Trim, ""))) & ", " & _
                                                "DRUD_REMARKS = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DRUD_REMARKS").ToString.Trim, ""))) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                              "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                              "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' " & _
                                              "and DRUD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("DRUD_SEQ").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                Dim maxLocDT As DataTable
                maxLocDT = gDB.getDataTable("select drud_loc, v_location.wh_code from wms_drum_dtl, (select max(drud_date) as max_drud_date from wms_drum_dtl " & _
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and DRUM_ID = '" & gU.dbEncode(ViewState("DRUM_ID")) & "') max_date, v_location where wms_drum_dtl.drud_date = max_date.max_drud_date " & _
                         "and wms_drum_dtl.imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                         "and wms_drum_dtl.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and wms_drum_dtl.DRUM_ID = '" & gU.dbEncode(ViewState("DRUM_ID")) & "' and v_location.loc = wms_drum_dtl.drud_loc and wms_drum_dtl.imp_code = v_location.imp_code order by drud_seq desc", gConn, transaction)

                If maxLocDT.Rows.Count > 0 Then
                    If emptyDrum_YN.Value <> "Y" Then

                        sql_string = "Update wms_drum set drum_loc='" & gU.dbEncode(maxLocDT.Rows(0).Item("drud_loc").ToString.Trim) & "', DRUM_WH_CODE='" & maxLocDT.Rows(0).Item("wh_code").ToString.Trim & "'" & _
                                     "where wms_drum.imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                     "and wms_drum.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                     "and wms_drum.DRUM_ID = '" & gU.dbEncode(ViewState("DRUM_ID")) & "'"
                        gDB.amendData(sql_string, gConn, transaction)

                        'If Session("pagemode") = "N" Then
                        '    sql_string = "SELECT WH_CODE, FL_NUM, AR_CODE, RK_CODE, BN_CODE FROM WMS_WH_BIN where LOC_KEY='" & maxLocDT.Rows(0).Item("drud_loc").ToString.Trim & "'"
                        '    Dim locDT As DataTable = gDB.getDataTable(sql_string, gConn, transaction)

                        '    If locDT.Rows.Count > 0 Then
                        '        sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                        '                     " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "') " & _
                        '                     " INSERT INTO WMS_ITEM_LOC_BAL " & _
                        '                     " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, ILOC_LOC, ILOC_WH, ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN, ILOC_BAL_QTY,  " & _
                        '                     " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, VND_CODE) " & _
                        '                     " VALUES ('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'000','" & maxLocDT.Rows(0).Item("drud_loc").ToString.Trim & "'," & _
                        '                     "'" & locDT.Rows(0).Item("WH_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("FL_NUM").ToString.Trim & "','" & locDT.Rows(0).Item("AR_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("RK_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("BN_CODE").ToString.Trim & "', 0,'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "') "
                        '        gDB.amendData(sql_string, gConn, transaction)

                        '        Dim tempIloc_seq As String = gU.decodeNullOrEmpty(DB.getValueFromSQL("select ILOC_SEQ from wms_item_loc_bal where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                        '                          "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ILOC_LOC = '" & gU.dbEncode(maxLocDT.Rows(0).Item("drud_loc").ToString.Trim) & "' " & _
                        '                          "and VND_CODE = '" & gU.dbEncode(ViewState("DRUM_ID")) & "'", gConn, transaction), "")

                        '        sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                        '                     " INSERT INTO WMS_ITEM_LOC_BAL_S " & _
                        '                     " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2, ILBS_ORG_QTY2, ILBS_ORG_SERIAL_NO,  " & _
                        '                     " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ILBS_SERIAL_NO, ILOC_SEQ) " & _
                        '                     " VALUES('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'" & gU.dbEncode(ViewState("DRUM_ID")) & "',0,'M',0,0,'" & gU.dbEncode(ViewState("DRUM_ID")) & "'," & _
                        '                     "'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "','" & tempIloc_seq & "') "
                        '        gDB.amendData(sql_string, gConn, transaction)
                        '    End If
                        'End If

                    Else

                        sql_string = "SELECT WH_CODE, FL_NUM, AR_CODE, RK_CODE, BN_CODE FROM WMS_WH_BIN where LOC_KEY='" & DRUM_LOC.Value & "'"
                        Dim locDT As DataTable = gDB.getDataTable(sql_string, gConn, transaction)

                        If locDT.Rows.Count > 0 Then

                            sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                         " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "') " & _
                                         " INSERT INTO WMS_ITEM_LOC_BAL " & _
                                         " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, ILOC_LOC, ILOC_WH, ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN, ILOC_BAL_QTY,  " & _
                                         " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, VND_CODE) " & _
                                         " VALUES ('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'000','" & DRUM_LOC.Value & "'," & _
                                         "'" & locDT.Rows(0).Item("WH_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("FL_NUM").ToString.Trim & "','" & locDT.Rows(0).Item("AR_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("RK_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("BN_CODE").ToString.Trim & "',0 ,'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "') "
                            gDB.amendData(sql_string, gConn, transaction)


                            Dim tempIloc_seq As String = gU.decodeNullOrEmpty(DB.getValueFromSQL("select ILOC_SEQ from wms_item_loc_bal where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ILOC_LOC = '" & gU.dbEncode(DRUM_LOC.Value) & "' " & _
                                              "and VND_CODE = '" & gU.dbEncode(ViewState("DRUM_ID")) & "'", gConn, transaction), "")

                            sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                         " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "') " & _
                                         " INSERT INTO WMS_ITEM_LOC_BAL_S " & _
                                         " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2, ILBS_ORG_QTY2, ILBS_ORG_SERIAL_NO,  " & _
                                         " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ILBS_SERIAL_NO, ILOC_SEQ) " & _
                                         " VALUES        ('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'" & gU.dbEncode(ViewState("DRUM_ID")) & "',0,'M',0,0,'" & gU.dbEncode(ViewState("DRUM_ID")) & "'," & _
                                         "'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "','" & tempIloc_seq & "') "
                            gDB.amendData(sql_string, gConn, transaction)

                            'sql_string = "Update WMS_ITEM_LOC_BAL set ILOC_WH='" & gU.dbEncode(DRUM_LOC.Value) & "', sys_lud=getdate(), sys_lub='" & Session("usr_id") & "''' " & _
                            '              " where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "'  and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "'" & _
                            '              " and itm_code= 'EMPTY_DRUM' " & _
                            '              " and VND_CODE = '" & gU.dbEncode(ViewState("DRUM_ID")) & "'"

                        End If
                    End If
                Else
                    If Session("pagemode") <> "N" Then

                        sql_string = "SELECT WH_CODE, FL_NUM, AR_CODE, RK_CODE, BN_CODE FROM WMS_WH_BIN where LOC_KEY='" & DRUM_LOC.Value & "'"
                        Dim locDT As DataTable = gDB.getDataTable(sql_string, gConn, transaction)

                        If locDT.Rows.Count > 0 Then

                            sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                         " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "') " & _
                                         " INSERT INTO WMS_ITEM_LOC_BAL " & _
                                         " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, ILOC_LOC, ILOC_WH, ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN, ILOC_BAL_QTY,  " & _
                                         " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, VND_CODE) " & _
                                         " VALUES ('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'000','" & DRUM_LOC.Value & "'," & _
                                         "'" & locDT.Rows(0).Item("WH_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("FL_NUM").ToString.Trim & "','" & locDT.Rows(0).Item("AR_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("RK_CODE").ToString.Trim & "','" & locDT.Rows(0).Item("BN_CODE").ToString.Trim & "',0 ,'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "') "
                            gDB.amendData(sql_string, gConn, transaction)


                            Dim tempIloc_seq As String = gU.decodeNullOrEmpty(DB.getValueFromSQL("select ILOC_SEQ from wms_item_loc_bal where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ILOC_LOC = '" & gU.dbEncode(DRUM_LOC.Value) & "' " & _
                                              "and VND_CODE = '" & gU.dbEncode(ViewState("DRUM_ID")) & "'", gConn, transaction), "")

                            sql_string = " if not exists (select 1 from WMS_ITEM_LOC_BAL_S where ILBS_DRUM_ID='" & gU.dbEncode(ViewState("DRUM_ID")) & "' and imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                         " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "') " & _
                                         " INSERT INTO WMS_ITEM_LOC_BAL_S " & _
                                         " (IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2, ILBS_ORG_QTY2, ILBS_ORG_SERIAL_NO,  " & _
                                         " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ILBS_SERIAL_NO, ILOC_SEQ) " & _
                                         " VALUES ('" & IMP_CODE.Value.Trim & "','" & STORER_CODE.SelectedValue & "','EMPTY_DRUM',1,'" & gU.dbEncode(ViewState("DRUM_ID")) & "',0,'M',0,0,'" & gU.dbEncode(ViewState("DRUM_ID")) & "'," & _
                                         "'" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "','" & gU.dbEncode(ViewState("DRUM_ID")) & "','" & tempIloc_seq & "') "
                            gDB.amendData(sql_string, gConn, transaction)
                        End If
                    End If
                End If


                transaction.Commit()

                successFlag = True
                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    DRUM_ID.Text = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    DRUM_ID.Enabled = False
                    'DRUM_CODE.CssClass = ""
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
        If ViewState("DRUM_ID") <> "" Then
            pk_code = ViewState("DRUM_ID")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("DRUM_ID"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("DRUM_ID") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        DRUM_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            DRUM_STATUS.Text = "ACTIVE"
            btnShowSL.Visible = False
            btnPrint.Visible = False
            btnPrnLabel.Visible = False

            Image_Loc_LookUp_DRUM_LOC.Visible = True
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT IMP_CODE, STORER_CODE, DRUM_ID, DRUM_TYPE, Convert(varchar, wms_drum.DRUM_DATE," & DDFORMAT & ") as DRUM_DATE, DRUM_WH_CODE, DRUM_LOC, DRUM_IN_WH, DRUM_CATEGORY, DRUM_DESC, DRUM_IS_RESERVED," & _
                        " DRUM_IS_INSP, DRUM_REMARKS, DRUM_SERIAL_NO,Convert(varchar, wms_drum.DRUM_CLOSE_DATE," & DDFORMAT & ") as DRUM_CLOSE_DATE, DRUM_REOPEN_BY, STATUS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB " & _
                        " FROM WMS_DRUM " & _
                        "where WMS_DRUM.DRUM_ID = '" & gU.dbEncode(pk_code) & "' " & _
                        "and WMS_DRUM.imp_code = '" & Session("IMP_CODE") & "' " & _
                        "and WMS_DRUM.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                DRUM_ID.Text = dt.Rows(0).Item("DRUM_ID").ToString

                DRUM_STATUS.Text = dt.Rows(0).Item("STATUS").ToString
                DRUM_TYPE.SelectedValue = dt.Rows(0).Item("DRUM_TYPE").ToString
                DRUM_LOC.Value = dt.Rows(0).Item("DRUM_LOC").ToString
                dsp_DRUM_LOC.Text = dt.Rows(0).Item("DRUM_LOC").ToString
                DRUM_DESC.Text = dt.Rows(0).Item("DRUM_DESC").ToString
                DRUM_DATE.Text = dt.Rows(0).Item("DRUM_DATE").ToString
                DRUM_WH_CODE.SelectedValue = dt.Rows(0).Item("DRUM_WH_CODE").ToString
                DRUM_WH_CODE.Enabled = False
                DRUM_REMARKS.Text = dt.Rows(0).Item("DRUM_REMARKS").ToString


                DRUM_REOPEN_BY.Text = dt.Rows(0).Item("DRUM_REOPEN_BY").ToString
                DRUM_IN_WH.SelectedValue = dt.Rows(0).Item("DRUM_IN_WH").ToString
                DRUM_SERIAL_NO.Text = dt.Rows(0).Item("DRUM_SERIAL_NO").ToString
                DRUM_CLOSE_DATE.Text = dt.Rows(0).Item("DRUM_CLOSE_DATE").ToString

                If dt.Rows(0).Item("DRUM_IS_RESERVED").ToString = "Y" Then DRUM_IS_RESERVED.Checked = True Else DRUM_IS_RESERVED.Checked = False
                If dt.Rows(0).Item("DRUM_IS_INSP").ToString = "Y" Then DRUM_IS_INSP.Checked = True Else DRUM_IS_INSP.Checked = False

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If DRUM_ID.Text <> "" Then
                    DRUM_ID.ReadOnly = True
                    DRUM_ID.BorderWidth = 0
                    DRUM_ID.BackColor = Drawing.Color.Transparent
                End If

                If dt.Rows(0).Item("STATUS").ToString = "INACTIVE" Then
                    btnReOpen.Visible = True
                End If

                If dt.Rows(0).Item("DRUM_TYPE").ToString = "DRUM" Then
                    btnShowSL.Visible = True
                    btnShowSL.Attributes.Add("onclick", "goToSL('" & storerCode & "','" & pk_code & "');return false;")


                    SQLString = "Select count(*) as count from WMS_ITEM_LOC_BAL_S where WMS_ITEM_LOC_BAL_S.imp_code = '" & Session("IMP_CODE") & "' and WMS_ITEM_LOC_BAL_S.storer_code = '" & gU.dbEncode(storerCode) & "' and upper(WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID)=upper('" & dt.Rows(0).Item("DRUM_ID").ToString & "')"

                    Dim balCount As Integer = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)

                    If balCount = 0 Then
                        Image_Loc_LookUp_DRUM_LOC.Visible = True
                        emptyDrum_YN.Value = "Y"
                        DRUM_WH_CODE.Enabled = True
                    Else
                        Image_Loc_LookUp_DRUM_LOC.Visible = False
                        emptyDrum_YN.Value = "N"
                        DRUM_WH_CODE.Enabled = False
                    End If

                Else
                    btnShowSL.Visible = False
                End If

                If dt.Rows(0).Item("DRUM_TYPE").ToString = "TROLLEY" Then
                    SQLString = " SELECT count(distinct itm_sku_no) as itm_count " & _
                                " FROM WMS_DELV_ORDER INNER JOIN " & _
                                " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND  " & _
                                " WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " where WMS_DELV_ORDER.DO_STATUS = 'PICKED'  " & _
                                " and WMS_DELV_ORDER.imp_code = '" & Session("IMP_CODE") & "' " & _
                                " and WMS_DELV_ORDER.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                                " and WMS_DELV_ORDER_d.DOD_TROLLEY_ID ='" & gU.dbEncode(pk_code) & "' "

                    TOTAL_PK_ITEM.Text = DB.getValueFromSQL(SQLString)

                End If

                btnPrint.Visible = True
                btnPrint.Attributes.Add("onclick", "goToPrint('" & storerCode & "','" & pk_code & "');return false;")
                btnPrnLabel.Visible = True

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = " SELECT 'U' as mFlag, WMS_DRUM_DTL.IMP_CODE, WMS_DRUM_DTL.STORER_CODE, WMS_DRUM_DTL.DRUM_ID, WMS_DRUM_DTL.DRUD_SEQ, DRUD_DATE as full_DRUD_DATE, convert(varchar,DRUD_DATE," & DDFORMAT & ") as DRUD_DATE, DRUD_MOVEMENT, DRUD_DOC_TYPE, DRUD_DOC_NO, DRUD_BY, DRUD_LOC, DRUD_MV_TYPE, " & _
                    " DRUD_REMARKS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " & _
                    " datepart(hh,drud_date) as DRUD_DATE_HR, datepart(mi,drud_date) as DRUD_DATE_MIN, " & _
                    " V_LOCATION.WH_CODE as DRUD_WH " & _
                    " FROM WMS_DRUM_DTL " & _
                    " LEFT OUTER JOIN V_LOCATION ON WMS_DRUM_DTL.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_DRUM_DTL.DRUD_LOC = V_LOCATION.LOC " & _
                    " where WMS_DRUM_DTL.DRUM_ID = '" & gU.dbEncode(pk_code) & "' " & _
                    " and WMS_DRUM_DTL.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    " and WMS_DRUM_DTL.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by WMS_DRUM_DTL.DRUD_DATE DESC, CONVERT(int, WMS_DRUM_DTL.DRUD_SEQ)"
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
        Dim cancelSql As String = "update WMS_DRUM " & _
                     "set status = 'INACTIVE', " & _
                     "DRUM_CLOSE_DATE = Getdate(), " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            DRUM_STATUS.Text = "INACTIVE"

            ar.sec_write = "N"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
            btnReOpen.Visible = True
            btnReOpen.Enabled = True
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

    Protected Sub btnCollect_Click(sender As Object, e As System.EventArgs) Handles btnCollect.Click
        Call addNewMove("C")
    End Sub

    Protected Sub btnReturn_Click(sender As Object, e As System.EventArgs) Handles btnReturn.Click
        Call addNewMove("T")
    End Sub

    Protected Sub btnRelocate_Click(sender As Object, e As System.EventArgs) Handles btnRelocate.Click
        Call addNewMove("L")
    End Sub

    Private Sub addNewMove(ByVal flag As String)

        Dim tempDT As DataTable

        tempDT = ViewState("dt")

        If tempDT IsNot Nothing Then

            If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then
                Dim rows_count As Integer = 0
                REM **********************
                REM Modify Here
                Dim seq_string As String = "select MAX(CAST(DRUD_SEQ AS int)) + 1 from WMS_DRUM_DTL " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' "
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

                tempDT.Rows.Add()
                rows_count = tempDT.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("DRUD_SEQ") = ViewState("n_cur_seq").ToString
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
                dt.Rows(rows_count - 1).Item("DRUD_MV_TYPE") = "MANUAL"
                dt.Rows(rows_count - 1).Item("DRUD_DATE") = Now.Date.ToString("dd/MM/yyyy")

                Select Case flag
                    Case "C"
                        dt.Rows(rows_count - 1).Item("DRUD_MOVEMENT") = "ISS"

                    Case "T"
                        dt.Rows(rows_count - 1).Item("DRUD_MOVEMENT") = "RTN"

                    Case "L"
                        dt.Rows(rows_count - 1).Item("DRUD_MOVEMENT") = "REL"

                End Select

                dt.AcceptChanges()

                ViewState("dt") = tempDT
                GridView1.DataSource = tempDT
                GridView1.DataBind()

            End If



        End If

    End Sub

    Protected Sub btnReOpen_Click(sender As Object, e As System.EventArgs) Handles btnReOpen.Click
        Dim cancelSql As String = "update WMS_DRUM " & _
                    "set status = 'ACTIVE', " & _
                    "DRUM_REOPEN_BY= '" & gU.dbEncode(Session("usr_id")) & "', " & _
                    "DRUM_CLOSE_DATE = NULL, " & _
                    "sys_lub = '" & Session("usr_id") & "', " & _
                    "sys_lud = Getdate() " & _
                    "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' "
        Dim successFlag As Boolean = False

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()
            successFlag = True
            DRUM_STATUS.Text = "INACTIVE"

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

        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "DRUMMAIN.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("DRUM_ID", DRUM_ID.Text.Trim)
            rmtPost.alertMsg = "This Drum/Trolley has been reopened!"
            rmtPost.Post()
        End If

    End Sub

    Function checkIsemptyDrum() As String
        Dim seq_string As String = "select 1 from WMS_ITEM_LOC_BAL_S " & _
                                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                    "and ILBS_DRUM_ID = '" & gU.dbEncode(DRUM_ID.Text.Trim) & "' " & _
                                                    "and ILBS_QTY2 > 0"
        REM **********************
        Dim nS_dt As New DataTable
        nS_dt = gDB.getDataTable(seq_string)


        If nS_dt.Rows.Count = 0 Then
            Return "Yes"
        Else
            Return "No"
        End If
    End Function

End Class
