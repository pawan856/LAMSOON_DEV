Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_SLBU_SLBalUpdtMain
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
    Private balDt As New DataTable
    Private DDFORMAT As String = "DD/MM/YYYY"
    Private exceptionEditList As List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")

        REM ****************************
        REM Modify Access Right Here
        'ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        'If ar.sessionExpired = "Y" Then
        '    ar.Force_PageEndCtrlClear(Me, False)
        '    Me.Visible = False
        '    Exit Sub
        'End If

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Short Length Cable Update Balance"

            'If Session("PAGE_SESSION_MENU_CODE") = "OP_SLBU" Then

            lbl_ImageHd.Text = "Enter New Balance"
            lbl_GV_Header.Text = "Existing Balance"

            lbl_STORER_CODE.Text = "Storer :"
            lbl_itm_sku_no.Text = "Stock No. :"
            lbl_pack_key.Text = "Pack Key :"
            
            btnSearch.Text = "Search"
            newrow.Text = "Add"
            btnPost.Text = "Post"

            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?"");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?"")){getLoad();}else{return false;}"
            If Session("pagemode") = "N" Then
                'TR_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "Short Length Cable Update Balance"

            lbl_ImageHd.Text = "Enter New Balance"
            lbl_GV_Header.Text = "Existing Balance"

            lbl_STORER_CODE.Text = "Storer :"
            lbl_itm_sku_no.Text = "Stock No. :"
            lbl_pack_key.Text = "Pack Key :"

            btnSearch.Text = "Search"
            newrow.Text = "新增"
            btnPost.Text = "發布"

            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        itm_sku_no.CssClass = "REQUIRED"
        pack_key.CssClass = "REQUIRED"
        STORER_CODE.CssClass = "REQUIRED"        
        REM **********************

        If Session("pagemode") = "N" Then
            'TR_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing
            ViewState("balDt") = Nothing
            newrow.Enabled = False
            'Call BindGV()
        Else
            dt = ViewState("dt")
            balDt = ViewState("balDt")
        End If

        If moduleAction = "SELECTIM" Then
            'addItemtoSTF()
            'dsp_TR_TO_LOC.Text = TR_TO_LOC.Value
            'ElseIf moduleAction = "SELECTTOWH" Then
            'changeToPallet()
        End If

        'setPageCtrlAccess()


        'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
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

    Private Sub BindGV()
        Dim SQLString As String = ""
        
        Dim sku_no As String = itm_sku_no.Text.Trim
        Dim packkey As String = pack_key.Text.Trim
        Dim storerCode As String = STORER_CODE.SelectedValue

        dt = New DataTable
        balDt = New DataTable

        ViewState("dt") = Nothing
        ViewState("balDt") = Nothing

        gvEdit.DataSource = Nothing
        gvEdit.DataBind()

        newrow.Enabled = True

        SQLString = ""
        SQLString += "select wms_item.imp_code, wms_item.storer_code, wms_item.itm_code, "
        SQLString += "wms_item.pack_key,  wms_item_loc_bal.iloc_pallet_no,  wms_item_loc_bal.iloc_loc, "
        SQLString += "wms_item_loc_bal.iloc_batch_no,  wms_item.itm_name,  wms_item.itm_sku_no, "
        SQLString += "wms_item_loc_bal_s.ilbs_serial_no,  wms_item_loc_bal_s.ilbs_uom2,  wms_item_loc_bal_s.ilbs_qty2, "
        SQLString += "wms_item_loc_bal_s.ilbs_drum_id, wms_item_loc_bal_s.ilbs_drum_level, wms_item_loc_bal.vnd_code, '' as mFlag "
        SQLString += "from wms_item_loc_bal inner join wms_item on "
        SQLString += "wms_item.imp_code = wms_item_loc_bal.imp_code "
        SQLString += "and wms_item.storer_code = wms_item_loc_bal.storer_code "
        SQLString += "and  wms_item.itm_code = wms_item_loc_bal.itm_code "
        SQLString += "and wms_item.pack_key = wms_item_loc_bal.pack_key "
        SQLString += "inner join wms_item_loc_bal_s on "
        SQLString += "wms_item_loc_bal.iloc_seq = wms_item_loc_bal_s.iloc_seq "
        SQLString += "where wms_item.itm_type = 'CABLE' "
        SQLString += "and wms_item.storer_code = '" & gU.dbEncode(storerCode) & "' "
        SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
        SQLString += "and wms_item.itm_sku_no = '" & gU.dbEncode(sku_no) & "' "
        SQLString += "and wms_item.pack_key = '" & gU.dbEncode(packkey) & "' "
        SQLString += "order by wms_item.itm_code, wms_item.pack_key, vnd_code "

        balDt = gDB.getDataTable(SQLString)

        If balDt.Rows.Count > 0 Then
            gvView.DataSource = balDt
            dt = balDt.Clone

            Dim newRow As DataRow

            newRow = dt.NewRow()
            newRow.ItemArray = balDt.Rows(0).ItemArray.Clone
            newRow.Item("ILOC_BATCH_NO") = DBNull.Value
            newRow.Item("ILBS_SERIAL_NO") = DBNull.Value
            newRow.Item("ILBS_QTY2") = DBNull.Value
            newRow.Item("ILOC_LOC") = DBNull.Value
            newRow.Item("ILBS_DRUM_ID") = DBNull.Value
            newRow.Item("ILBS_DRUM_LEVEL") = DBNull.Value
            newRow.Item("mFlag") = "N"

            dt.Rows.Add(newRow)
            dt.AcceptChanges()

            gvEdit.DataSource = dt
            gvEdit.DataBind()
        Else
            gvView.DataSource = Nothing

            SQLString = ""
            SQLString += "select wms_item.imp_code, wms_item.storer_code, wms_item.itm_code, wms_item.itm_sku_no, "
            SQLString += "wms_item.pack_key,  wms_item.itm_name, "
            SQLString += "wms_item_loc_bal.iloc_pallet_no,  wms_item_loc_bal.iloc_loc,  wms_item_loc_bal.iloc_batch_no, "
            SQLString += "wms_item_loc_bal_s.ilbs_serial_no,  wms_item.itm_uom2 as ilbs_uom2,  wms_item_loc_bal_s.ilbs_qty2, "
            SQLString += "wms_item_loc_bal_s.ilbs_drum_id, wms_item_loc_bal_s.ilbs_drum_level, wms_alt_vend_item.vnd_code, 'N' as mFlag "
            SQLString += "from wms_item left outer join wms_alt_vend_item on "
            SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            SQLString += "left outer join wms_item_loc_bal on "
            SQLString += "wms_item.imp_code = wms_item_loc_bal.imp_code "
            SQLString += "and wms_item.storer_code = wms_item_loc_bal.storer_code "
            SQLString += "and  wms_item.itm_code = wms_item_loc_bal.itm_code "
            SQLString += "and wms_item.pack_key = wms_item_loc_bal.pack_key "
            SQLString += "left outer join wms_item_loc_bal_s on "
            SQLString += "wms_item_loc_bal.iloc_seq = wms_item_loc_bal_s.iloc_seq "
            SQLString += "where wms_item.itm_type = 'CABLE' "
            SQLString += "and wms_item.storer_code = '" & gU.dbEncode(storerCode) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_sku_no = '" & gU.dbEncode(sku_no) & "' "
            SQLString += "and wms_item.pack_key = '" & gU.dbEncode(packkey) & "' "
            SQLString += "order by wms_item.itm_code, wms_item.pack_key, vnd_code "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then
                'balDt = itmRowDt
                'dt = balDt.Clone
                'addNewRow()
                gvEdit.DataSource = dt
                gvEdit.DataBind()
            Else
                newrow.Enabled = False
            End If
        End If

        ViewState("dt") = dt
        ViewState("balDt") = balDt
        gvView.DataBind()
    End Sub

    Private Sub addNewRow()
        If cU.gfBuildDataTableforGridView(dt, gvEdit, True) Then            
            Dim newRow As DataRow

            If dt.Rows.Count > 0 Then
                newRow = dt.NewRow()
                newRow.ItemArray = dt.Rows(0).ItemArray.Clone
                newRow.Item("ILOC_BATCH_NO") = DBNull.Value
                newRow.Item("ILBS_SERIAL_NO") = DBNull.Value
                newRow.Item("ILBS_QTY2") = DBNull.Value
                newRow.Item("ILOC_LOC") = DBNull.Value
                newRow.Item("ILBS_DRUM_ID") = DBNull.Value
                newRow.Item("ILBS_DRUM_LEVEL") = DBNull.Value
                newRow.Item("mFlag") = "N"

                dt.Rows.Add(newRow)
                dt.AcceptChanges()
            Else
                If balDt.Rows.Count = 0 Then
                    Dim sku_no As String = itm_sku_no.Text.Trim
                    Dim packkey As String = pack_key.Text.Trim
                    Dim storerCode As String = STORER_CODE.SelectedValue

                    Dim SQLString As String = ""
                    SQLString += "select wms_item.imp_code, wms_item.storer_code, wms_item.itm_code, wms_item.itm_sku_no, "
                    SQLString += "wms_item.pack_key,  wms_item.itm_name, "
                    SQLString += "wms_item_loc_bal.iloc_pallet_no,  wms_item_loc_bal.iloc_loc,  wms_item_loc_bal.iloc_batch_no, "
                    SQLString += "wms_item_loc_bal_s.ilbs_serial_no,  wms_item.itm_uom2 as ilbs_uom2,  wms_item_loc_bal_s.ilbs_qty2, "
                    SQLString += "wms_item_loc_bal_s.ilbs_drum_id, wms_item_loc_bal_s.ilbs_drum_level, wms_alt_vend_item.vnd_code, 'N' as mFlag "
                    SQLString += "from wms_item left outer join wms_alt_vend_item on "
                    SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
                    SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
                    SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
                    SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
                    SQLString += "left outer join wms_item_loc_bal on "
                    SQLString += "wms_item.imp_code = wms_item_loc_bal.imp_code "
                    SQLString += "and wms_item.storer_code = wms_item_loc_bal.storer_code "
                    SQLString += "and  wms_item.itm_code = wms_item_loc_bal.itm_code "
                    SQLString += "and wms_item.pack_key = wms_item_loc_bal.pack_key "
                    SQLString += "left outer join wms_item_loc_bal_s on "
                    SQLString += "wms_item_loc_bal.iloc_seq = wms_item_loc_bal_s.iloc_seq "
                    SQLString += "where wms_item.itm_type = 'CABLE' "
                    SQLString += "and wms_item.storer_code = '" & gU.dbEncode(storerCode) & "' "
                    SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
                    SQLString += "and wms_item.itm_sku_no = '" & gU.dbEncode(sku_no) & "' "
                    SQLString += "and wms_item.pack_key = '" & gU.dbEncode(packkey) & "' "
                    SQLString += "order by wms_item.itm_code, wms_item.pack_key, vnd_code "

                    dt = gDB.getDataTable(SQLString)
                Else
                    newRow = dt.NewRow()
                    newRow.ItemArray = balDt.Rows(0).ItemArray.Clone
                    newRow.Item("ILOC_BATCH_NO") = DBNull.Value
                    newRow.Item("ILBS_SERIAL_NO") = DBNull.Value
                    newRow.Item("ILBS_QTY2") = DBNull.Value
                    newRow.Item("ILOC_LOC") = DBNull.Value
                    newRow.Item("ILBS_DRUM_ID") = DBNull.Value
                    newRow.Item("ILBS_DRUM_LEVEL") = DBNull.Value
                    newRow.Item("mFlag") = "N"

                    dt.Rows.Add(newRow)
                    dt.AcceptChanges()
                End If
            End If        

            ViewState("dt") = dt
            gvEdit.DataSource = dt
            gvEdit.DataBind()
        End If
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        BindGV()
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        addNewRow()
    End Sub

    Protected Sub gvView_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvView.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim
                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim

                CType(e.Row.FindControl("ILOC_BATCH_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("ILBS_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("ILBS_QTY2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_QTY2").ToString.Trim                
                CType(e.Row.FindControl("ILBS_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_UOM2").ToString.Trim
                CType(e.Row.FindControl("ILOC_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_LOC").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_ID"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_ID").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_LEVEL"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_LEVEL").ToString.Trim

        End Select
    End Sub

    Protected Sub gvEdit_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvEdit.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim
                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim

                CType(e.Row.FindControl("ILOC_BATCH_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("ILBS_SERIAL_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("ILBS_QTY2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_QTY2").ToString.Trim
                CType(e.Row.FindControl("ILBS_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_UOM2").ToString.Trim

                CType(e.Row.FindControl("dsp_ILOC_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_LOC").ToString.Trim
                CType(e.Row.FindControl("ILOC_LOC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ILOC_LOC").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_ID"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_ID").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_LEVEL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_LEVEL").ToString.Trim

                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                nImage.Attributes.Add("onclick", "LocLookUp('" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ILOC_LOC"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ILOC_LOC"), HiddenField).ClientID) & "')")

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If
        End Select
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvEdit.RowDeleting                
        dt.Rows.RemoveAt(e.RowIndex)        
        dt.AcceptChanges()

        ViewState("dt") = dt
        gvEdit.DataSource = dt
        gvEdit.DataBind()
    End Sub

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click        
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        If gvEdit.Rows.Count > 0 Then
            If cU.gfBuildDataTableforGridView(dt, gvEdit, True) Then
                If validateAll() Then
                    Try
                        For Each rows As DataRow In dt.Rows
                            Dim whCode As String = DB.getValueFromSQL("select wh_code from v_location where loc = '" & gU.dbEncode(rows.Item("iloc_loc").ToString) & "'")
                            'serial_No = wFun.getSerialNo(STORER_CODE.SelectedValue, rows.Item("rtd_itm_code").ToString.Trim, rows.Item("rtd_pack_key").ToString.Trim, rows.Item("RTD_SERIAL_NO").ToString.Trim)

                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNullOrEmpty(rows.Item("itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNullOrEmpty(rows.Item("pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "SLBU"
                            st.IO_DOC_ID = ""
                            st.IO_QTY = 1
                            st.IO_CBM = 0
                            st.IO_KG = 0
                            st.PALLET_NO = "000"
                            st.IO_WH = whCode
                            st.IO_LOC = gU.decodeNullOrEmpty(rows.Item("iloc_loc").ToString.Trim, "")
                            st.IO_EXPIRY_DATE = ""
                            st.IO_MANU_DATE = ""
                            st.lO_BATCH_NO = gU.decodeNullOrEmpty(rows.Item("iloc_batch_no").ToString.Trim, "")
                            st.lO_VND_CODE = gU.decodeNullOrEmpty(rows.Item("vnd_code").ToString.Trim, "")

                            st.IOS_DRUM_ID = gU.decodeNullOrEmpty(rows.Item("ilbs_drum_id").ToString.Trim, "")
                            st.IOS_DRUM_LEVEL = gU.decodeNullOrEmpty(rows.Item("ilbs_drum_level").ToString.Trim, "")

                            st.IOS_SERIAL_NO = gU.decodeNullOrEmpty(rows.Item("ilbs_serial_no").ToString.Trim, "")

                            st.IOS_QTY2 = gU.decodeEmptyCdbl(rows.Item("ilbs_qty2").ToString.Trim, 0)
                            st.IOS_UOM2 = gU.decodeNullOrEmpty(rows.Item("ilbs_uom2").ToString.Trim, "")
                            st.IOS_SL = "Y"

                            Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ilbs_serial_no").ToString.Trim, ""), gConn, transaction)

                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)

                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        Next

                        transaction.Commit()

                        'btnPost.Enabled = False

                        uiFun.displayMsgNew(updtGVEdit, "1007", "", Session("gLang"))

                        If gConn IsNot Nothing Then
                            If gConn.State = ConnectionState.Open Then
                                gConn.Close()
                                gConn.Dispose()
                            End If
                        End If

                    Catch ex As Exception
                        transaction.Rollback()
                        gConn.Close()
                        Throw ex
                    End Try

                    'Call Windows Close Script
                    BindGV()

                End If
            End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtGVEdit, "", "No item can be posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtGVEdit, "", "没有可供发布的物件!", Session("gLang"))
            End If
        End If

        load_ModalPopupExtender.Hide()
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        Dim itemCount As Integer = 0

        Dim gotExists As Boolean = False

        If gvEdit.Rows.Count > 0 Then
            For i = 0 To gvEdit.Rows.Count - 1
                If dt.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                    If CType(gvEdit.Rows(i).FindControl("ilbs_serial_no"), TextBox).Text.Trim = "" Then
                        If Session("gLang") = "E" Then                            
                            uiFun.displayMsgNew(updtGVEdit, "", "Serial No. cannot be empty!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtGVEdit, "", "Serial No.不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(gvEdit.Rows(i).FindControl("ilbs_qty2"), TextBox).Text.Trim = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtGVEdit, "", "Qty2 cannot be empty!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtGVEdit, "", "數量2不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(gvEdit.Rows(i).FindControl("ilbs_qty2"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtGVEdit, "", "Invalid number, Qty2!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtGVEdit, "", "無效的數字, 數量2!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CLng(CType(gvEdit.Rows(i).FindControl("ilbs_qty2"), TextBox).Text.Trim) = 0 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtGVEdit, "", "Qty2 cannot be zero!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtGVEdit, "", "數量2不能設零!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(gvEdit.Rows(i).FindControl("iloc_loc"), HiddenField).Value.Trim = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtGVEdit, "", "Please select Location!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtGVEdit, "", "位置不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

                    itemCount += 1
                End If
            Next

            For i = 0 To gvEdit.Rows.Count - 1
                If balDt IsNot Nothing AndAlso balDt.Rows.Count > 0 Then
                    Dim dRows As DataRow() = balDt.Select("iloc_loc = '" & gU.dbEncode(CType(gvEdit.Rows(i).FindControl("iloc_loc"), HiddenField).Value.Trim) & "' " & _
                                                          "and isnull(iloc_batch_no,'') = '" & gU.dbEncode(CType(gvEdit.Rows(i).FindControl("iloc_batch_no"), TextBox).Text.Trim) & "' " & _
                                                          "and isnull(ilbs_serial_no,'') = '" & gU.dbEncode(CType(gvEdit.Rows(i).FindControl("ilbs_serial_no"), TextBox).Text.Trim) & "' " & _
                                                          "and isnull(ilbs_drum_id,'') = '" & gU.dbEncode(CType(gvEdit.Rows(i).FindControl("ilbs_drum_id"), TextBox).Text.Trim) & "' " & _
                                                          "and isnull(ilbs_drum_level,'') = '" & gU.dbEncode(CType(gvEdit.Rows(i).FindControl("ilbs_drum_level"), TextBox).Text.Trim) & "' ")

                    If dRows.Count > 0 Then
                        For x As Integer = 0 To gvEdit.Rows(i).Cells.Count - 1
                            gvEdit.Rows(i).Cells(x).BackColor = Drawing.Color.LightCoral
                        Next
                        gotExists = True
                    Else
                        For x As Integer = 0 To gvEdit.Rows(i).Cells.Count - 1
                            gvEdit.Rows(i).Cells(x).BackColor = Nothing
                        Next
                    End If

                    CType(gvEdit.Rows(i).FindControl("dsp_iloc_loc"), Label).Text = CType(gvEdit.Rows(i).FindControl("iloc_loc"), HiddenField).Value
                End If
            Next
        End If

        If itemCount = 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtGVEdit, "", "Please Add Item!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtGVEdit, "", "Please Add Item!", Session("gLang"))
            End If
            Return False
        End If

        If gotExists Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtGVEdit, "", "Records already exists in balance!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtGVEdit, "", "Records already exists in balance!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function
End Class
