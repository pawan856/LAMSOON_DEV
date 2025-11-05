Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Globalization

Partial Class OUTBOUND_COMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private dl As New DocLink
    Private moduleAction As String = ""
    Private DDFORMAT2 As String = "dd/MM/yyyy"
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private imp_code As String = ""
    Private ModuleAbb As String = "CO_MAIN"
    Private dt As New DataTable
    Private dtUOM, dtUOM2 As DataTable
    Private exceptionEditList As List(Of String)
    Private gvCol() As String = {"COD_DISP_SEQ", "BTNSPLIT", "COD_WH_CODE", "COD_REQ_BY", "COD_PLANT", "COD_TICKET_NO", "COD_CARTON_NO", "COD_PALLET_NO", "COD_BATCH_NO", "COD_LOT_NO",
                                 "COD_VND_CODE", "COD_ITM_CODE", "ITM_SKU_NO", "COD_ITM_DESC", "COD_EXPIRY_DATE", "COD_MANU_DATE", "COD_PACKING", "COD_PACK_KEY", "COD_QTY", "COD_OS_QTY",
                                 "COD_PCS_CARTON", "COD_CARTON_PALLET", "COD_UOM", "COD_QTY2", "COD_UOM2", "COD_PCS_UOM", "COD_TOTPCS", "COD_TOT_WGT", "COD_TOT_CBM",
                                 "COD_REM", "COD_BATCH_NO", "COD_BATCH_NO"}
    Dim TotalUnitPrice As Double = 0.0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT2 = gU.getConfig("DDFORMAT2")
        'DDFORMAT = gU.getConfig("DDFORMAT")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_CO", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("STO_BATCH_FIELD_REF") = ""

            uiFun.load_dropdown(CO_SHIP_MODE, "select colc_code, colc_eng_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))

            If Session("pagemode") = "N" Then
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            End If
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            'btnConfirm.Visible = False
            btnClose.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")

                If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
                    STORER_CODE.Enabled = False
                End If
                Call setDefStorerInfo()
            End If
            If CO_DATE.Text = "" Then
                CO_DATE.Text = Now.Date.ToString(DDFORMAT2)
            End If
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Customer Order Maintenance"
            lbl_ImageHd.Text = "Customer Order Items"
            lbl_CO_CODE.Text = "CO Code:"
            lbl_CO_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_CO_CUS_REF_NO.Text = "EBS SO No.:"
            lbl_CO_DATE.Text = "Order Date:"
            lbl_PRJ_CODE.Text = "SO Header ID:"
            lbl_CUS_CODE.Text = "Customer:"
            lbl_co_addr.Text = "Address:"
            lbl_CO_AREA_DEL.Text = "Area:"
            lbl_CO_DELIVER_TO.Text = "Deliver To:"
            lbl_CO_REGION_DEL.Text = "Region:"
            lbl_CO_COUNTRY_DEL.Text = "Country:"
            lbl_co_cus_cont.Text = "Attention:"
            lbl_co_cus_cont_tel.Text = "Tel. No.:"
            lbl_CO_REM.Text = "Remarks:"
            'lbl_CO_TRACK_NO.Text = "Tracking No.:"
            lbl_CO_FTRACK_NO.Text = "Forwarder Tracking No.:"
            lbl_CO_INV_NO.Text = "Invoice No.:"
            lbl_ROUTE_ID.Text = "ROUTE"
            lbl_CO_PROVINCE.Text = "Province:"
            lbl_CO_CITY.Text = "City"
            lbl_CO_SENDER.Text = "Sender Name"
            lbl_CO_SENDER_COUNTRY.Text = "Sender Country"
            lbl_CO_SENDER_PROVINCE.Text = "Sender Province"
            lbl_CO_SENDER_REGION.Text = "Sender Region"
            lbl_CO_SENDER_ADDR.Text = "Sender Address"
            lbl_CO_SENDER_TEL.Text = "Sender Telephone"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"

            btnNOTE.Text = "Notes"
            'btnConfirm.Text = "Close"
            btnClose.Text = "Close"
            reOpenBtn.Text = "Re-Open"

            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this order?"");"
            reOpenBtn.OnClientClick = "return confirm(""Are you sure to Re-open this order?"");"
            'btnConfirm.OnClientClick = "return confirm(""Are you sure to confirm this order?"");"
            btnClose.OnClientClick = "return confirm(""Are you sure to close this order?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this order?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this order?"");"

            If Session("pagemode") = "N" Then
                CO_CODE.Text = "[No. will be auto generated]"
            End If
        ElseIf Session("gLang") = "C" Then
            lheader.Value = "客戶訂單維護"
            lbl_ImageHd.Text = "客戶訂單"
            lbl_CO_CODE.Text = "系統訂單號碼:"
            lbl_CO_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_CO_CUS_REF_NO.Text = "EBS SO號:"
            lbl_CO_DATE.Text = "訂單日期:"
            lbl_PRJ_CODE.Text = "SO Header ID:"
            lbl_CUS_CODE.Text = "客户:"
            lbl_co_addr.Text = "地址:"
            lbl_CO_AREA_DEL.Text = "區域:"
            lbl_CO_DELIVER_TO.Text = "運送到:"
            lbl_CO_REGION_DEL.Text = "地區:"
            lbl_CO_COUNTRY_DEL.Text = "國家:"
            lbl_co_cus_cont.Text = "聯絡人:"
            lbl_co_cus_cont_tel.Text = "聯絡電話:"
            lbl_CO_REM.Text = "備註:"
            'lbl_CO_TRACK_NO.Text = "追查編號:"
            lbl_CO_FTRACK_NO.Text = "運送追查編號:"
            lbl_CO_INV_NO.Text = "发票編號:"
            lbl_ROUTE_ID.Text = "路線"
            lbl_CO_CUST_INV_NO.Text = "客戶发票編號:"
            lbl_CO_TARGET_DELDATE.Text = "目標運送日期:"
            lbl_CO_SHIP_MODE.Text = "運送模式:"
            lbl_CO_SENDER.Text = "寄件人姓名"
            lbl_CO_SENDER_COUNTRY.Text = "寄件人國家"
            lbl_CO_SENDER_PROVINCE.Text = "寄件人省份"
            lbl_CO_SENDER_REGION.Text = "寄件人地區"
            lbl_CO_SENDER_ADDR.Text = "寄件人地址"
            lbl_CO_SENDER_TEL.Text = "寄件人電話"

            lbl_CO_PROVINCE.Text = "省:"
            lbl_CO_CITY.Text = "市:"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"

            btnTrack.Text = "訂單追蹤"
            btnNOTE.Text = "工單"
            'btnConfirm.Text = "確認"
            btnClose.Text = "關閉"
            reOpenBtn.Text = "Re-Open"
            reOpenBtn.OnClientClick = "return confirm(""Are you sure to Re-open this order?"");"
            CancelBtn.OnClientClick = "return confirm(""確定取消訂單?"");"
            'btnConfirm.OnClientClick = "return confirm(""確定確認訂單?"");"
            btnClose.OnClientClick = "return confirm(""確定關閉訂單?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存訂單?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存訂單?"");"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPrtLbl.Text = "打印物品標籤"
            cSBBtn.Text = "檢查貨品存庫"
            selectItemBtn.Text = "選擇物品"
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        CO_DATE.CssClass = "REQUIRED"
        ROUTE_ID.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            STORER_CODE.CssClass = "REQUIRED"
        End If

        If Not IsPostBack Then
            'ViewState("dt") = Nothing

            gU.clearSessionTempData(ModuleAbb)

            ViewState("n_cur_seq") = ""
            ViewState("CO_CODE") = ""
            ViewState("FrmUP") = ""

            ViewState("CO_CODE") = Server.UrlDecode(Request("co_code"))
            ViewState("STORER_CODE") = Server.UrlDecode(Request("storer_code"))
            ViewState("FrmUP") = Server.UrlDecode(Request("FrmUP"))

            Call BindGV()
        Else
            'dt = ViewState("dt")
            dt = gU.getSessionTempData(ModuleAbb, "codDT", Nothing)
        End If

        'Set Access Right
        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            If Session("usr_pref_storer") <> STORER_CODE.Text Then
                Response.End()
            End If
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoCO()
        End If

        Call changeLabel()

        Dim colIdx_StartWith As Integer = 1
        'ar.addColDef("COD_SEQ", "cod_seq", colIdx_StartWith)
        ar.addColDef("COD_DISP_SEQ", "cod_disp_seq", colIdx_StartWith)
        'ar.addColDef("COD_JOB_NO", "cod_job_no", colIdx_StartWith)
        ar.addColDef("COD_WH_CODE", "COD_WH_CODE", colIdx_StartWith)
        ar.addColDef("COD_REQ_BY", "COD_REQ_BY", colIdx_StartWith)
        ar.addColDef("COD_PLANT", "cod_plant", colIdx_StartWith)
        ar.addColDef("COD_TICKET_NO", "COD_TICKET_NO", colIdx_StartWith)
        'ar.addColDef("COD_CUST_CODE", "cod_cust_code", colIdx_StartWith)
        'ar.addColDef("COD_ITM_PARENT", "cod_itm_parent", colIdx_StartWith)
        ar.addColDef("COD_CARTON_NO", "cod_carton_no", colIdx_StartWith)
        ar.addColDef("COD_PALLET_NO", "cod_pallet_no", colIdx_StartWith)
        ar.addColDef("COD_BATCH_NO", "cod_batch_no", colIdx_StartWith)
        ar.addColDef("COD_LOT_NO", "cod_lot_no", colIdx_StartWith)
        ar.addColDef("COD_VND_CODE", "cod_vnd_code", colIdx_StartWith)
        ar.addColDef("COD_ITM_CODE", "cod_itm_code", colIdx_StartWith)
        ar.addColDef("ITM_SKU_NO", "itm_sku_no", colIdx_StartWith)
        ar.addColDef("COD_ITM_DESC", "cod_itm_desc", colIdx_StartWith)
        'ar.addColDef("ITM_DESC", "itm_desc", colIdx_StartWith)
        ar.addColDef("COD_EXPIRY_DATE", "cod_expiry_date", colIdx_StartWith)
        ar.addColDef("COD_MANU_DATE", "cod_manu_date", colIdx_StartWith)
        ar.addColDef("COD_PACKING", "cod_packing", colIdx_StartWith)
        ar.addColDef("COD_PACK_KEY", "cod_pack_key", colIdx_StartWith)
        ar.addColDef("COD_QTY", "cod_qty", colIdx_StartWith)
        ar.addColDef("COD_OS_QTY", "COD_OS_QTY", colIdx_StartWith)
        ar.addColDef("COD_PCS_CARTON", "cod_pcs_carton", colIdx_StartWith)
        ar.addColDef("COD_CARTON_PALLET", "cod_carton_pallet", colIdx_StartWith)
        ar.addColDef("COD_UOM", "cod_uom", colIdx_StartWith)
        ar.addColDef("COD_QTY2", "cod_qty2", colIdx_StartWith)
        ar.addColDef("COD_UOM2", "cod_uom2", colIdx_StartWith)
        ar.addColDef("COD_PCS_UOM", "cod_pcs_uom", colIdx_StartWith)
        ar.addColDef("COD_TOTPCS", "cod_totpcs", colIdx_StartWith)
        ar.addColDef("COD_TOT_WGT", "cod_tot_wgt", colIdx_StartWith)
        ar.addColDef("COD_TOT_CBM", "cod_tot_cbm", colIdx_StartWith)
        ar.addColDef("COD_REM", "cod_rem", colIdx_StartWith)
        ar.addColDef("HOLD_QTY", "hold_qty", colIdx_StartWith)

        cm = New CommonMenu("CO", lheader.Value, CO_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "N"
        cm.haveAttachments = "N"
        cm.haveNotes = "N"
        cm.haveTasks = "N"
        cm.haveEmail = "N"
        cm.haveHistory = "N"

        cm.genCM(cmBar)

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "CO", CO_CODE.Text, "../../")

        REM Select RO button
        setPageCtrlAccess()

        If CO_STATUS.Text = "CANCELLED" Then
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');return false;")
        Else
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');return false;")
        End If

        If CO_DATA_FR.Value = "CSMS" Then
            ar.sec_viewMode = "Y"
            If CO_STATUS.Text <> "CANCELLED" OrElse CO_STATUS.Text <> "CLOSED" Then
                exceptionEditList.Add("CancelBtn")
                exceptionEditList.Add("btnClose")
            End If
        End If

        If CO_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            'btnConfirm.Visible = False
            btnClose.Visible = False
        ElseIf CO_STATUS.Text = "CLOSED" Then
            'ar.sec_write = "N"
            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
            'btnConfirm.Visible = False
            btnClose.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
        'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "CO", "WMS_CUST_ORDER_D")

        btnNOTE.Enabled = True

        If CO_STATUS.Text = "CLOSED" Then
            reOpenBtn.Visible = True
            reOpenBtn.Enabled = True
        Else
            reOpenBtn.Visible = False
            reOpenBtn.Enabled = False
        End If

        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            btnPrtLbl.Visible = False
        End If


        'ar.setFieldCustomize(Me, "OB_CO", STORER_CODE.SelectedValue, "WMS_CUST_ORDER")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "OB_CO", STORER_CODE.SelectedValue, "WMS_CUST_ORDER_D", gvCol)
        'End If

        Dim qtyIndex As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For v As Integer = 0 To GridView1.Columns.Count - 1
                If GridView1.Columns(v).Visible Then
                    If GridView1.Columns(v).AccessibleHeaderText = "cod_qty" Then
                        Exit For
                    Else
                        qtyIndex += 1
                    End If
                End If
            Next
        End If

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OB_CO','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("CO_CODE") & "','N');")

        ViewState("qtyIndex") = qtyIndex

        Dim coCode As String = Server.UrlDecode(Request("co_code"))
        Dim SQLString As String = ""
        SQLString = "Select co_status FROM  wms_cust_order WHERE wms_cust_order.imp_code='" & Session("IMP_CODE") & "' and wms_cust_order.STORER_CODE='" & ViewState("STORER_CODE") & "' " &
                                                         "AND wms_cust_order.co_code = '" & gU.dbEncode(coCode) & "' "

        Dim coCodeValue As DataTable = gDB.getDataTable(SQLString)
        If (coCodeValue.Rows.Count > 0) Then

            If coCodeValue.Rows(0).Item(0).ToString().Trim <> "NEW" Then
                CancelBtn.Enabled = False
            Else
                CancelBtn.Enabled = True
            End If
        Else
            CancelBtn.Enabled = True
        End If
    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)
        exceptionEditList.Add("btnPrtLbl")
        exceptionEditList.Add("btnTrack")

    End Sub

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Then
            CType(ctl, Button).Enabled = True
            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        Dim selectSql As String

        selectSql = "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1"

        dtUOM = gDB.getDataTable(selectSql)

        selectSql = "select colc_code as code, colc_eng_value as name from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_UOM2' order by COLC_DISPLAY_SEQ"

        dtUOM2 = gDB.getDataTable(selectSql)
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        'If e.CommandName = "SplitItem" Then
        Select Case e.CommandName
            Case "SplitItem"
                If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                    Dim rows_count As Integer = 0
                    Dim rowNum As Integer = -1
                    Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                    If Not IsNothing(gvRow) Then
                        rowNum = gvRow.RowIndex
                    End If

                    Dim seq_string As String = "select MAX(convert(int, COD_SEQ)) + 1 from WMS_CUST_ORDER_D " &
                                       "where IMP_CODE = '" & gU.dbEncode(imp_code) & "' " &
                                       "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                       "and CO_CODE = '" & gU.dbEncode(CO_CODE.Text.Trim) & "' "

                    Dim nS_dt As New DataTable
                    nS_dt = gDB.getDataTable(seq_string)
                    Dim next_seq_no As String
                    Dim temp_seq_no As Integer = 1

                    next_seq_no = ""

                    If Session("pagemode") <> "N" Then
                        If nS_dt.Rows.Count > 0 Then
                            next_seq_no = nS_dt.Rows(0).Item(0).ToString()
                        Else
                            next_seq_no = "0"
                        End If
                    Else
                        next_seq_no = rowNum + 1
                    End If

                    If ViewState("n_cur_seq") = "" Then
                        ViewState("n_cur_seq") = next_seq_no
                    Else
                        temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                        ViewState("n_cur_seq") = temp_seq_no.ToString
                    End If


                    Dim newRow As DataRow

                    newRow = dt.NewRow
                    dt.Rows.InsertAt(newRow, rowNum + 1)

                    'dt.Rows.Add()
                    rows_count = dt.Rows.Count

                    dt.Rows(rowNum + 1).Item("cod_seq") = ViewState("n_cur_seq").ToString
                    'dt.Rows(rowNum + 1).Item("GRD_DISP_SEQ") = ViewState("n_cur_seq").ToString
                    dt.Rows(rowNum + 1).Item("COD_DISP_SEQ") = dt.Rows(rowNum).Item("COD_DISP_SEQ")
                    dt.Rows(rowNum + 1).Item("COD_REF_SEQ") = CInt(dt.Rows(rowNum).Item("cod_seq").ToString)

                    dt.Rows(rowNum + 1).Item("COD_JOB_NO") = dt.Rows(rowNum).Item("COD_JOB_NO")
                    dt.Rows(rowNum + 1).Item("COD_TICKET_NO") = dt.Rows(rowNum).Item("COD_TICKET_NO")
                    dt.Rows(rowNum + 1).Item("COD_CUST_CODE") = dt.Rows(rowNum).Item("COD_CUST_CODE")
                    dt.Rows(rowNum + 1).Item("COD_CUST_REM") = dt.Rows(rowNum).Item("COD_CUST_REM")
                    dt.Rows(rowNum + 1).Item("COD_ITM_PARENT") = dt.Rows(rowNum).Item("COD_ITM_PARENT")
                    dt.Rows(rowNum + 1).Item("COD_ITM_CODE") = dt.Rows(rowNum).Item("COD_ITM_CODE")
                    dt.Rows(rowNum + 1).Item("ITM_SKU_NO") = dt.Rows(rowNum).Item("itm_sku_no")
                    dt.Rows(rowNum + 1).Item("COD_PACK_KEY") = dt.Rows(rowNum).Item("COD_PACK_KEY")
                    dt.Rows(rowNum + 1).Item("COD_PALLET_NO") = dt.Rows(rowNum).Item("COD_PALLET_NO")
                    dt.Rows(rowNum + 1).Item("COD_ITM_DESC") = dt.Rows(rowNum).Item("COD_ITM_DESC")
                    'dt.Rows(rowNum + 1).Item("ITM_DESC") = dt.Rows(rowNum).Item("ITM_DESC")
                    dt.Rows(rowNum + 1).Item("COD_PACKING") = dt.Rows(rowNum).Item("COD_PACKING")
                    dt.Rows(rowNum + 1).Item("COD_UOM") = dt.Rows(rowNum).Item("COD_UOM")
                    dt.Rows(rowNum + 1).Item("COD_QTY2") = dt.Rows(rowNum).Item("COD_QTY2")
                    dt.Rows(rowNum + 1).Item("COD_UOM2") = dt.Rows(rowNum).Item("COD_UOM2")
                    dt.Rows(rowNum + 1).Item("COD_PCS_UOM") = dt.Rows(rowNum).Item("COD_PCS_UOM")
                    dt.Rows(rowNum + 1).Item("COD_DELIV_QTY") = dt.Rows(rowNum).Item("COD_DELIV_QTY")
                    dt.Rows(rowNum + 1).Item("COD_TOT_WGT") = dt.Rows(rowNum).Item("COD_TOT_WGT")
                    dt.Rows(rowNum + 1).Item("COD_TOT_CBM") = dt.Rows(rowNum).Item("COD_TOT_CBM")
                    dt.Rows(rowNum + 1).Item("COD_VND_CODE") = dt.Rows(rowNum).Item("COD_VND_CODE")
                    dt.Rows(rowNum + 1).Item("COD_LOT_NO") = dt.Rows(rowNum).Item("COD_LOT_NO")
                    dt.Rows(rowNum + 1).Item("COD_CARTON_NO") = dt.Rows(rowNum).Item("COD_CARTON_NO")
                    dt.Rows(rowNum + 1).Item("COD_WH_CODE") = dt.Rows(rowNum).Item("COD_WH_CODE")
                    dt.Rows(rowNum + 1).Item("COD_REQ_BY") = dt.Rows(rowNum).Item("COD_REQ_BY")
                    dt.Rows(rowNum + 1).Item("COD_PLANT") = dt.Rows(rowNum).Item("COD_PLANT")
                    dt.Rows(rowNum + 1).Item("COD_BATCH_NO") = dt.Rows(rowNum).Item("COD_BATCH_NO")
                    dt.Rows(rowNum + 1).Item("COD_EXPIRY_DATE") = dt.Rows(rowNum).Item("COD_EXPIRY_DATE")
                    dt.Rows(rowNum + 1).Item("COD_MANU_DATE") = dt.Rows(rowNum).Item("COD_MANU_DATE")
                    dt.Rows(rowNum + 1).Item("COD_QTY") = 0
                    dt.Rows(rowNum + 1).Item("COD_TOTPCS") = 0
                    dt.Rows(rowNum + 1).Item("HOLD_QTY") = 0

                    dt.Rows(rowNum + 1).Item("mFlag") = "N"

                    dt.AcceptChanges()

                    gU.setSessionTempData(ModuleAbb, "codDT", dt)

                    GridView1.DataSource = dt
                    GridView1.DataBind()
                End If


            Case "HOLD"

                ' gU.setSessionTempData(ModuleAbb, "codDT", ViewState("dt"))



        End Select

        'End If
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                e.Row.Cells(0).ID = "COD_SEQ"

                CType(e.Row.FindControl("cod_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_disp_seq").ToString.Trim
                'CType(e.Row.FindControl("cod_job_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_job_no").ToString.Trim
                CType(e.Row.FindControl("COD_TICKET_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_TICKET_NO").ToString.Trim

                'CType(e.Row.FindControl("cod_cust_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_cust_code").ToString.Trim
                'CType(e.Row.FindControl("cod_itm_parent"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_itm_parent").ToString.Trim

                'CType(e.Row.FindControl("cod_vnd_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_vnd_code").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("cod_vnd_code"), DropDownList),
                                    "select distinct vnd_code " &
                                    "from wms_alt_vend_item " &
                                    "where " &
                                    "wms_alt_vend_item.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                    "and wms_alt_vend_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and wms_alt_vend_item.itm_code = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim) & "' " &
                                    "and wms_alt_vend_item.pack_key = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim) & "' " &
                                    "order by vnd_code ", "vnd_code", "vnd_code", , Session("gSelectLabel"))

                CType(e.Row.FindControl("cod_vnd_code"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_vnd_code").ToString.Trim

                CType(e.Row.FindControl("h_cod_itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim
                CType(e.Row.FindControl("h_cod_pack_key"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim

                CType(e.Row.FindControl("cod_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("cod_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim
                CType(e.Row.FindControl("cod_itm_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_itm_desc").ToString.Trim
                'CType(e.Row.FindControl("itm_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_desc").ToString.Trim
                CType(e.Row.FindControl("cod_packing"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_packing").ToString.Trim
                CType(e.Row.FindControl("cod_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "cod_qty").ToString.Trim)
                CType(e.Row.FindControl("cod_os_qty"), Label).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "cod_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "cod_post_qty").ToString.Trim, 0)

                GetUnitPrice(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "cod_qty").ToString.Trim))

                'Dim nDropDown As DropDownList = CType(e.Row.FindControl("cod_uom"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_uom").ToString.Trim
                'CType(e.Row.FindControl("cod_uom"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_uom").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("cod_uom"), DropDownList), dtUOM, "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "COD_UOM").ToString.Trim)

                uiFun.load_dropdown(CType(e.Row.FindControl("cod_uom2"), DropDownList), dtUOM2, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "COD_UOM2").ToString.Trim)

                CType(e.Row.FindControl("cod_qty2"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "cod_qty2").ToString.Trim)

                'Dim pDropDown As DropDownList = CType(e.Row.FindControl("cod_pallet_no"), DropDownList)

                'uiFun.load_dropdown(pDropDown, "select distinct iloc_pallet_no from wms_item_loc_bal where " & _
                '                                "imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND " & _
                '                                "STORER_CODE = '" & STORER_CODE.SelectedValue.Trim & "' AND " & _
                '                                "ITM_CODE = '" & DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim & "' AND " & _
                '                                "PACK_KEY = '" & DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim & "' " & _
                '                                " order by 1", "iloc_pallet_no", "iloc_pallet_no", "000", "000")

                'pDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_pallet_no").ToString.Trim

                'CType(e.Row.FindControl("cod_pallet_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_pallet_no").ToString.Trim

                CType(e.Row.FindControl("cod_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_pallet_no").ToString.Trim

                CType(e.Row.FindControl("cod_pcs_uom"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "cod_pcs_uom").ToString.Trim)
                CType(e.Row.FindControl("cod_totpcs"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "cod_totpcs").ToString.Trim)
                CType(e.Row.FindControl("cod_tot_wgt"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "cod_tot_wgt").ToString.Trim)

                If IsDBNull(DataBinder.Eval(e.Row.DataItem, "cod_tot_cbm")) Then
                    CType(e.Row.FindControl("cod_tot_cbm"), TextBox).Text = "0"
                Else
                    CType(e.Row.FindControl("cod_tot_cbm"), TextBox).Text = CDbl(DataBinder.Eval(e.Row.DataItem, "cod_tot_cbm")).ToString("###,###,##0.0000")
                End If

                CType(e.Row.FindControl("cod_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "cod_rem").ToString.Trim

                CType(e.Row.FindControl("HOLD_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim
                CType(e.Row.FindControl("HOLD_QTY"), TextBox).Attributes.Add("readonly", "readonly")
                REM **********************

                CType(e.Row.FindControl("cod_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("cod_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("cod_pcs_uom"), TextBox).ClientID & ".value * this.value")
                CType(e.Row.FindControl("cod_qty"), TextBox).Attributes("onkeyup") = "sumqtytotal(this.value)"
                CType(e.Row.FindControl("cod_pcs_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("cod_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("cod_qty"), TextBox).ClientID & ".value * this.value")


                'uiFun.load_dropdown(CType(e.Row.FindControl("cod_batch_no"), DropDownList), _
                '                    "select distinct dc_date_code " & _
                '                    "from wms_item d, wms_item_loc_bal l, wms_date_code dc " & _
                '                    "where " & _
                '                    "d.imp_code = l.imp_code " & _
                '                    "and d.storer_code = l.storer_code " & _
                '                    "and d.itm_code = l.itm_code " & _
                '                    "and d.pack_key = l.pack_key " & _
                '                    "and l.iloc_batch_no = dc.dc_date_code " & _
                '                    "and l.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                '                    "and l.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                '                    "and l.itm_code = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim) & "' " & _
                '                    "and l.pack_key = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim) & "' " & _
                '                    "order by dc_date_code desc ", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))

                uiFun.load_dropdown(CType(e.Row.FindControl("cod_batch_no"), DropDownList), "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                CType(e.Row.FindControl("cod_batch_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "cod_batch_no").ToString.Trim
                CType(e.Row.FindControl("ori_batch"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ori_batch").ToString.Trim


                CType(e.Row.FindControl("COD_LOT_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "COD_EXPIRY_DATE").ToString.Trim
                If DataBinder.Eval(e.Row.DataItem, "COD_EXPIRY_DATE").ToString.Trim = "" Then
                    CType(e.Row.FindControl("COD_LOT_NO"), Label).Text = ""
                Else
                    Dim d As DateTime = DateTime.ParseExact(CType(e.Row.FindControl("COD_LOT_NO"), Label).Text, "dd/MM/yyyy", CultureInfo.InvariantCulture)
                    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    CType(e.Row.FindControl("COD_LOT_NO"), Label).Text = reformatted
                End If


                CType(e.Row.FindControl("COD_CARTON_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_CARTON_NO").ToString.Trim
                CType(e.Row.FindControl("COD_PLANT"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_PLANT").ToString.Trim
                CType(e.Row.FindControl("COD_WH_CODE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_WH_CODE").ToString.Trim
                CType(e.Row.FindControl("COD_REQ_BY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_REQ_BY").ToString.Trim

                CType(e.Row.FindControl("COD_PCS_CARTON"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "COD_PCS_CARTON").ToString.Trim)
                CType(e.Row.FindControl("COD_CARTON_PALLET"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "COD_CARTON_PALLET").ToString.Trim)

                CType(e.Row.FindControl("COD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("COD_MANU_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COD_MANU_DATE").ToString.Trim

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                'nButton.ID = "btnDelete" & e.Row.RowIndex

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                CType(e.Row.FindControl("btnHold"), Button).OnClientClick = "javascript:HoldStock('" & DataBinder.Eval(e.Row.DataItem, "storer_code").ToString.Trim & "', " &
                                                                            "'" & DataBinder.Eval(e.Row.DataItem, "co_code").ToString.Trim & "', " &
                                                                            "'" & DataBinder.Eval(e.Row.DataItem, "cod_seq").ToString.Trim & "'," &
                                                                            "'" & CType(e.Row.FindControl("HOLD_QTY"), TextBox).ClientID & "'," &
                                                                            "'" & CType(e.Row.FindControl("cod_batch_no"), DropDownList).ClientID & "'" &
                                                                            ");"
                'If DataBinder.Eval(e.Row.DataItem, "cod_batch_no").ToString.Trim <> "" Then
                '    CType(e.Row.FindControl("btnHold"), Button).Enabled = True
                'Else
                '    CType(e.Row.FindControl("btnHold"), Button).Enabled = False
                'End If

                If DataBinder.Eval(e.Row.DataItem, "COD_REF_SEQ").ToString.Trim <> "" Then
                    DirectCast(e.Row.FindControl("btnSplit"), Button).Visible = False
                    DirectCast(e.Row.FindControl("cod_disp_seq"), TextBox).Visible = False

                    If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then

                        For k As Integer = 0 To e.Row.Cells.Count - 1
                            e.Row.Cells(k).CssClass = "GV_SPLIT"
                        Next
                    End If
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                    CType(e.Row.FindControl("btnHold"), Button).Enabled = False
                    'CType(e.Row.FindControl("HOLD_QTY"), TextBox).Visible = False
                End If


                If ar.hasBtnRight("BT_CO_HOLD") Then
                    'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim <> "N" Then
                    '    If DataBinder.Eval(e.Row.DataItem, "cod_batch_no").ToString.Trim <> "" Then CType(e.Row.FindControl("btnHold"), Button).Enabled = True
                    '    CType(e.Row.FindControl("HOLD_QTY"), TextBox).Visible = True
                    'End If
                Else
                    CType(e.Row.FindControl("btnHold"), Button).Enabled = False
                    CType(e.Row.FindControl("HOLD_QTY"), TextBox).Visible = False
                End If

                ar.hideGVForStorer(GridView1, STORER_CODE.Text, "CO", "WMS_CUST_ORDER_D", e)
            Case DataControlRowType.Footer
                CType(e.Row.FindControl("cod_total_qty"), Label).Text = GetTotal()
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select max(convert(int,COD_SEQ)) + 1 from wms_cust_order_d " &
                                        "where imp_code = '" & gU.dbEncode(imp_code.Trim.ToString) & "' " &
                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and co_code = '" & gU.dbEncode(CO_CODE.Text) & "' "
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
            dt.Rows(rows_count - 1).Item("cod_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            'ViewState("dt") = dt
            gU.setSessionTempData(ModuleAbb, "codDT", dt)

            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        Dim tempDT As New DataTable
        Dim holdDT As New DataTable

        tempDT = gU.getSessionTempData(ModuleAbb, "codDT", Nothing)
        holdDT = gU.getSessionTempData(ModuleAbb, "holdDT", Nothing)

        'If CO_CODE.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_CO_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_CO_CODE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If CO_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_CO_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_CO_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(CO_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_CO_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_CO_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If ROUTE_ID.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_ROUTE_ID.Text & " cannot be empty! Please enter the value of ROUTE!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_ROUTE_ID.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1

                If uiFun.gvValidate(Me, dt, "cod_qty", "Qty",
                                 CType(GridView1.Rows(i).FindControl("cod_qty"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "cod_pcs_uom", "Number/ UOM",
                                 CType(GridView1.Rows(i).FindControl("cod_pcs_uom"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "cod_totpcs", "Total Number",
                                 CType(GridView1.Rows(i).FindControl("cod_totpcs"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "cod_tot_wgt", "Weight(kg)",
                                 CType(GridView1.Rows(i).FindControl("cod_tot_wgt"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "cod_tot_cbm", "CBM",
                                 CType(GridView1.Rows(i).FindControl("cod_tot_cbm"), TextBox).Text) = False Then Return False


                If gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("HOLD_QTY"), TextBox).Text, 0) >
                   gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("cod_qty"), TextBox).Text, 0) Then

                    uiFun.displayMsg(Me, "", "Hold Quantity cannot larger than CO Quantity!", Session("gLang"))
                    Return False

                End If
            Next
        End If

        Dim SQLString As String = ""
        Dim total_Qty As Integer = 0
        Dim HSTK_QTY As Integer = 0
        Dim RO_QTY As Integer = 0
        Dim HRO_QTY As Integer = 0

        Dim c_code As String = ViewState("CO_CODE")
        Dim s_code As String = ViewState("STORER_CODE")

        If tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                If tempDT.Rows(i).Item("mFlag").ToString.Trim = "U" Then
                    HSTK_QTY = 0

                    If gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("HOLD_QTY"), TextBox).Text, 0) > 0 Then
                        'SQLString = " Select (LOC_QTY - Hold_QTY) as total_qty from  " & _
                        '  " (SELECT ISNULL(SUM(ILOC_BAL_QTY),0) as LOC_QTY " & _
                        '  "  FROM wms_item_loc_bal " & _
                        '  "  WHERE wms_item_loc_bal.IMP_CODE= '" & gU.dbEncode(Session("imp_code")) & "' AND wms_item_loc_bal.STORER_CODE='" & gU.dbEncode(s_code) & "' AND wms_item_loc_bal.ITM_CODE='" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " & _
                        '  "  AND wms_item_loc_bal.PACK_KEY= '" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' AND wms_item_loc_bal.ILOC_PALLET_NO=ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "','000') AND wms_item_loc_bal.ILOC_BATCH_NO='" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue) & "') LOC,  " & _
                        '  " (SELECT ISNULL(SUM(COH_QTY),0) as Hold_QTY " & _
                        '  " FROM wms_cust_order, wms_cust_order_d, wms_cust_order_hold " & _
                        '  "  WHERE wms_cust_order.IMP_CODE     = wms_cust_order_d.IMP_CODE " & _
                        '  "  AND wms_cust_order.STORER_CODE = wms_cust_order_d.STORER_CODE " & _
                        '  "  AND wms_cust_order.CO_CODE     = wms_cust_order_d.CO_CODE " & _
                        '  "  AND wms_cust_order.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        '  "  AND wms_cust_order_d.IMP_CODE     = wms_cust_order_hold.IMP_CODE " & _
                        '  "  AND wms_cust_order_d.STORER_CODE = wms_cust_order_hold.STORER_CODE  " & _
                        '  "  AND wms_cust_order_d.CO_CODE     = wms_cust_order_hold.CO_CODE  " & _
                        '  "  AND wms_cust_order_d.COD_SEQ     = wms_cust_order_hold.COD_SEQ  " & _
                        '  "  AND wms_cust_order_hold.COH_STATUS  = 'HOLD' AND wms_cust_order_hold.COH_TYPE  = 'STOCK' " & _
                        '  "  AND wms_cust_order.CO_CODE <> '" & gU.dbEncode(c_code) & "' " & _
                        '  "AND wms_cust_order_d.cod_seq <> '" & gU.dbEncode(tempDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "' " & _
                        '  "  AND wms_cust_order_d.IMP_CODE= '" & gU.dbEncode(Session("imp_code")) & "' " & _
                        '  "AND wms_cust_order_d.STORER_CODE='" & gU.dbEncode(s_code) & "' " & _
                        '  "AND wms_cust_order_d.COD_ITM_CODE='" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " & _
                        '  "  AND wms_cust_order_d.COD_PACK_KEY= '" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' " & _
                        '  "AND ISNULL(wms_cust_order_d.COD_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "','000') " & _
                        '  "AND wms_cust_order_d.COD_BATCH_NO='" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue) & "') HOLD "

                        'total_Qty = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)


                        SQLString = " SELECT wms_replenish.ro_code,  wms_replenish_d.rod_seq,  wms_storer.sto_name,  wms_replenish_d.rod_qty - gr.gr_qty  as rod_qty, " &
                                    "        Convert(varchar, rod_manu_date," & DDFORMAT & ")   AS rod_manu_date,  Convert(varchar, rod_expiry_date," & DDFORMAT & ") AS rod_expiry_date " &
                                    " FROM wms_replenish,  wms_replenish_d, wms_storer, ( " &
                                    " SELECT ISNULL(Sum(grd_po_qty),0) as GR_QTY FROM wms_replenish, wms_replenish_d, wms_goodsrcv_d, wms_goodsrcv " &
                                    " WHERE wms_replenish.IMP_CODE     = wms_replenish_d.IMP_CODE AND wms_replenish.STORER_CODE = wms_replenish_d.STORER_CODE " &
                                    " AND wms_replenish.RO_CODE     = wms_replenish_d.RO_CODE AND wms_goodsrcv_d.IMP_CODE     = wms_replenish_d.IMP_CODE " &
                                    " AND wms_goodsrcv_d.STORER_CODE = wms_replenish_d.STORER_CODE AND wms_replenish_d.ROD_ITM_CODE   = wms_goodsrcv_d.GRD_ITM_CODE " &
                                    " AND wms_replenish_d.ROD_PACK_KEY  = wms_goodsrcv_d.GRD_PACK_KEY AND wms_replenish_d.ROD_PALLET_NO = wms_goodsrcv_d.GRD_PALLET_NO " &
                                    " AND wms_replenish_d.ROD_BATCH_NO  = wms_goodsrcv_d.GRD_BATCH_NO AND wms_goodsrcv_d.IMP_CODE     = wms_goodsrcv.IMP_CODE " &
                                    " AND wms_goodsrcv_d.STORER_CODE = wms_goodsrcv.STORER_CODE AND wms_goodsrcv_d.GR_CODE     = wms_goodsrcv.GR_CODE " &
                                    " AND wms_replenish.ro_status                 <> 'CLOSED' " &
                                    " AND wms_replenish.imp_code                   ='" & Session("IMP_CODE") & "' " &
                                    " AND wms_replenish.STORER_CODE                ='" & gU.dbEncode(s_code) & "' " &
                                    " AND wms_replenish_d.rod_itm_code             ='" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " &
                                    " AND wms_replenish_d.rod_batch_no             ='" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue) & "' " &
                                    " AND wms_replenish_d.rod_pack_key             ='" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' " &
                                    " AND ISNULL(wms_replenish_d.ROD_PALLET_NO,'000') =ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "','000')  " &
                                    " AND wms_goodsrcv.gr_status = 'POSTED' ) gr " &
                                    " WHERE wms_replenish.IMP_CODE                 = wms_replenish_d.IMP_CODE AND wms_replenish.STORER_CODE = wms_replenish_d.STORER_CODE " &
                                    " AND wms_replenish.RO_CODE                    = wms_replenish_d.RO_CODE AND wms_replenish.imp_code = wms_storer.imp_code " &
                                    " AND wms_replenish.storer_code = wms_storer.storer_code " &
                                    " AND wms_replenish_d.rod_qty                  > 0 " &
                                    " AND wms_replenish.ro_status                 <> 'CLOSED' " &
                                    " AND wms_replenish.imp_code                   ='" & Session("IMP_CODE") & "' " &
                                    " AND wms_replenish.STORER_CODE                ='" & gU.dbEncode(s_code) & "' " &
                                    " AND wms_replenish_d.rod_itm_code             ='" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " &
                                    " AND wms_replenish_d.rod_batch_no             ='" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue) & "' " &
                                    " AND wms_replenish_d.rod_pack_key             ='" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' " &
                                    " AND ISNULL(wms_replenish_d.ROD_PALLET_NO,'000') =ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "','000')  "

                        RO_QTY = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)

                        If Not holdDT Is Nothing AndAlso holdDT.Rows.Count > 0 Then
                            For x = 0 To holdDT.Rows.Count - 1
                                If holdDT.Rows(x).Item("mFlag").ToString.Trim <> "D" Then
                                    If holdDT.Rows(x).Item("cod_seq") = tempDT.Rows(i).Item("COD_SEQ").ToString.Trim Then

                                        If CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue <> "" AndAlso CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue <> holdDT.Rows(x).Item("COH_BATCH_NO").ToString.Trim Then
                                            uiFun.displayMsg(Me, "", "Batch No. not matched in hold record!", Session("gLang"))
                                            Return False
                                        End If

                                        If holdDT.Rows(x).Item("COH_STATUS").ToString.Trim = "HOLD" Then
                                            If holdDT.Rows(x).Item("COH_TYPE").ToString.Trim = "STOCK" Then

                                                HSTK_QTY += holdDT.Rows(x).Item("COH_QTY").ToString.Trim



                                                SQLString = "select LOC.LOC_QTY - isnull(HOLD.Hold_QTY, 0) AS VALUE " &
                                                            "FROM ( " &
                                                                "select sum(ILOC_BAL_QTY) as LOC_QTY " &
                                                                "from wms_item_loc_bal l " &
                                                                "where l.IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "' " &
                                                                "AND l.STORER_CODE = '" & gU.dbEncode(s_code) & "' " &
                                                                "AND l.ITM_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " &
                                                                "AND l.PACK_KEY = '" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' " &
                                                                "AND l.ILOC_PALLET_NO = ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "', '000') " &
                                                                "AND isnull(l.ILOC_BATCH_NO, '') = isnull('" & gU.dbEncode(holdDT.Rows(x).Item("COH_BATCH_NO").ToString.Trim) & "', '') " &
                                                                "and l.ILOC_BAL_QTY > 0) LOC " &
                                                            "left outer join ( " &
                                                                "select SUM(COH_QTY) as Hold_QTY " &
                                                                "from wms_cust_order c, wms_cust_order_d d, wms_cust_order_hold h " &
                                                                "WHERE c.IMP_CODE = d.IMP_CODE " &
                                                                "AND c.STORER_CODE = d.STORER_CODE " &
                                                                "AND c.CO_CODE = d.CO_CODE " &
                                                                "AND c.CO_STATUS <> 'CANCELLED' " &
                                                                "AND d.IMP_CODE = h.IMP_CODE " &
                                                                "AND d.STORER_CODE = h.STORER_CODE " &
                                                                "AND d.CO_CODE = h.CO_CODE " &
                                                                "AND d.COD_SEQ = h.COD_SEQ " &
                                                                "AND h.COH_STATUS = 'HOLD' " &
                                                                "AND h.COH_TYPE = 'STOCK' " &
                                                                "AND h.IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "' " &
                                                                "AND h.STORER_CODE = '" & gU.dbEncode(s_code) & "' " &
                                                                "AND h.COH_ITM_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("cod_itm_code").ToString.Trim) & "' " &
                                                                "AND h.COH_PACK_KEY = '" & gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim) & "' " &
                                                                "AND ISNULL(h.COH_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "', '000') " &
                                                                "AND isnull(h.COH_BATCH_NO, '') = isnull('" & gU.dbEncode(holdDT.Rows(x).Item("COH_BATCH_NO").ToString.Trim) & "', '') " &
                                                                "and h.COH_QTY > 0 " &
                                                                ") HOLD " &
                                                            "on 1=1 "

                                                '"AND ISNULL(h.COH_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim) & "', '000') " & _
                                                '"AND isnull(h.COH_BATCH_NO, '') = isnull('" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue) & "', '') " & _

                                                If CDbl(gU.decodeNullOrEmpty(holdDT.Rows(x).Item("COH_QTY").ToString.Trim, "0")) > CDbl(gU.decodeNullOrEmpty(DB.getValueFromSQL(SQLString), "0")) Then
                                                    uiFun.displayMsg(Me, "", "Hold Stock Quantity cannot larger than Available Qty! " & tempDT.Rows(i).Item("cod_itm_code").ToString.Trim & " - " & tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim & " - " & CType(GridView1.Rows(i).FindControl("COD_BATCH_NO"), DropDownList).SelectedValue, Session("gLang"))
                                                    Return False
                                                End If

                                            ElseIf holdDT.Rows(x).Item("COH_TYPE").ToString.Trim = "RO" Then


                                                SQLString = "Select ISNULL(SUM(COH_QTY),0) as hold_qty " &
                                                            "FROM wms_cust_order_hold, wms_cust_order " &
                                                            "WHERE wms_cust_order_hold.imp_code = wms_cust_order.imp_code " &
                                                            "AND wms_cust_order_hold.storer_code = wms_cust_order.storer_code " &
                                                            "AND wms_cust_order_hold.co_code = wms_cust_order.co_code " &
                                                            "AND wms_cust_order_hold.imp_code='" & Session("IMP_CODE") & "' " &
                                                            "AND wms_cust_order_hold.storer_code='" & gU.dbEncode(s_code) & "'" &
                                                            "AND wms_cust_order_hold.co_code <> '" & gU.dbEncode(c_code) & "' " &
                                                            "AND wms_cust_order_hold.cod_seq <> '" & gU.dbEncode(tempDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "' " &
                                                            "and coh_status='HOLD' and coh_type='RO' " &
                                                            "AND ro_code='" & gU.dbEncode(holdDT.Rows(x).Item("RO_CODE").ToString.Trim) & "' " &
                                                            "and ROD_SEQ='" & gU.dbEncode(holdDT.Rows(x).Item("ROD_SEQ").ToString.Trim) & "' " &
                                                            "AND wms_cust_order.co_status not in ('CLOSED', 'CANCELLED') "

                                                HRO_QTY = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)


                                                If holdDT.Rows(x).Item("COH_QTY").ToString.Trim > (RO_QTY - HRO_QTY) Then
                                                    uiFun.displayMsg(Me, "", "Hold Stock Quantity Cannot Larger than Available Qty.!", Session("gLang"))
                                                    Return False
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If


                        'If HSTK_QTY > total_Qty Then
                        '    uiFun.displayMsg(Me, "", "Hold Stock Quantity Cannot Larger than Stock Balance!", Session("gLang"))
                        '    Return False
                        'End If
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
        Dim insertSql As String
        Dim dupSQL As String
        Dim dupTbl As New DataTable

        Dim cod_batch_no As String

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("CO", gConn, transaction)

                    'nextNo = CO_CODE.Text
                    REM **********************

                    'If CO_TRACK_NO.Text = "" Then
                    '    CO_TRACK_NO.Text = DB.getDocNo("TRACKNO", gConn, transaction)
                    'End If

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    dupSQL = "select 1 from wms_cust_order " &
                            "where co_code = '" & gU.dbEncode(nextNo) & "' " &
                            "and imp_code='" & gU.dbEncode(imp_code) & "' " &
                            "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_cust_order ( " &
                                        "co_code, imp_code, storer_code, " &
                                        "co_status, co_cus_ref_no, co_project_no, " &
                                        "co_date, cus_code, cus_name, " &
                                        "co_addr1, co_addr2, co_addr3," &
                                        "co_area_del, CO_DELIVER_TO, co_region_del, co_country_del, " &
                                        "co_cus_cont, co_cus_cont_tel, co_rem, co_ftrack_no, " &
                                        "co_inv_no,route_id, co_cust_inv_no,CO_SHIP_MODE,CO_TARGET_DELDATE,CO_EDI_SIR_NO," &
                                        "CO_PROVINCE, CO_CITY, CO_SENDER, CO_SENDER_COUNTRY, CO_SENDER_PROVINCE, CO_SENDER_REGION, CO_SENDER_ADDR, CO_SENDER_TEL," &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values ( " &
                                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_STATUS.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_CUS_REF_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_PROJECT_NO.Text)) & ", " &
                                        gU.convdbDate(gU.dbEncode(CO_DATE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(CO_ADDR1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_ADDR2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_ADDR3.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(CO_AREA_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_DELIVER_TO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_REGION_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_COUNTRY_DEL.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(CO_CUS_CONT.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_CUS_CONT_TEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_REM.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(CO_FTRACK_NO.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(CO_INV_NO.Text)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(ROUTE_ID.Text)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_CUST_INV_NO.Text)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_SHIP_MODE.SelectedValue)) & "," &
                                        gU.convdbDate(gU.dbEncode(CO_TARGET_DELDATE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CO_EDI_SIR_NO.Text.Trim)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_PROVINCE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CO_CITY.Text.Trim)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_SENDER.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CO_SENDER_COUNTRY.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CO_SENDER_PROVINCE.Text.Trim)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(CO_SENDER_REGION.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CO_SENDER_ADDR.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CO_SENDER_TEL.Text.Trim)) & "," &
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                        REM **********************
                        'co_track_no,
                        ' gU.convdbNVCData(gU.dbEncode(CO_TRACK_NO.Text)) & ","

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            'uiFun.reOrderDetails(dt, "cod_disp_seq")
                            reOrderCOD(dt)

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                cod_batch_no = ""
                                cod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("cod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("COD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("COD_EXPIRY_DATE").ToString.Trim, ""))

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        If cod_batch_no <> "" Then
                                            dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(cod_batch_no) & "'"
                                            dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                            If dupTbl.Rows.Count <= 0 Then

                                                insertSql = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                            " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                            "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                            "'" & gU.dbEncode(cod_batch_no) & "'," &
                                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                                gDB.amendData(insertSql, gConn, transaction)
                                            End If
                                        End If

                                        itemSQL = "insert into wms_cust_order_d ( " &
                                                    "co_code, imp_code, storer_code, cod_seq, cod_disp_seq, " &
                                                    "cod_job_no, COD_TICKET_NO, cod_cust_code, " &
                                                    "cod_itm_parent, cod_itm_code, cod_pack_key, cod_itm_desc, " &
                                                    "cod_packing, cod_qty, cod_uom, cod_qty2, cod_uom2, cod_pcs_uom, " &
                                                    "cod_totpcs, cod_tot_wgt, cod_tot_cbm, cod_rem, " &
                                                    "cod_pallet_no, cod_vnd_code, cod_ref_seq, cod_batch_no, " &
                                                    "cod_lot_no, cod_carton_no, cod_plant, COD_WH_CODE, COD_REQ_BY, cod_pcs_carton, cod_carton_pallet, " &
                                                    "cod_expiry_date, cod_manu_date, " &
                                                    "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                    "values ( " &
                                                    gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_disp_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_job_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_TICKET_NO").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_cust_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_parent").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_pack_key").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_desc").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_packing").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty2").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom2").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_uom").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_totpcs").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_wgt").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_cbm").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_rem").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pallet_no").ToString.Trim, "000"))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_vnd_code").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_ref_seq").ToString.Trim, "null")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(cod_batch_no)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_lot_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_carton_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_plant").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_WH_CODE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_REQ_BY").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_carton").ToString.Trim, "null")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_carton_pallet").ToString.Trim, "null")) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                        ' gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_batch_no").ToString.Trim, ""))) & ", " & _
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next

                            If Session("usr_type") = "T" Or Session("usr_type") = "C" Then
                                SendCustCOMail(nextNo, STORER_CODE.SelectedValue)
                            End If


                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in " & lheader.Value & "!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", lheader.Value & "发现重复的资料录!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    nextNo = CO_CODE.Text
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_cust_order set " &
                                    "co_cus_ref_no = " & gU.convdbNVCData(gU.dbEncode(CO_CUS_REF_NO.Text)) & ", " &
                                    "co_project_no = " & gU.convdbNVCData(gU.dbEncode(CO_PROJECT_NO.Text)) & ", " &
                                    "co_status = " & gU.convdbNVCData(gU.dbEncode(CO_STATUS.Text)) & ", " &
                                    "co_date = " & gU.convdbDate(gU.dbEncode(CO_DATE.Text)) & ", " &
                                    "cus_code = " & gU.convdbNVCData(gU.dbEncode(CUS_CODE.Text)) & ", " &
                                    "cus_name = " & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text)) & ", " &
                                    "co_addr1 = " & gU.convdbNVCData(gU.dbEncode(CO_ADDR1.Text)) & ", " &
                                    "co_addr2 = " & gU.convdbNVCData(gU.dbEncode(CO_ADDR2.Text)) & ", " &
                                    "co_addr3 = " & gU.convdbNVCData(gU.dbEncode(CO_ADDR3.Text)) & ", " &
                                    "co_area_del = " & gU.convdbNVCData(gU.dbEncode(CO_AREA_DEL.Text)) & ", " &
                                    "CO_DELIVER_TO = " & gU.convdbNVCData(gU.dbEncode(CO_DELIVER_TO.Text)) & ", " &
                                    "co_region_del = " & gU.convdbNVCData(gU.dbEncode(CO_REGION_DEL.Text)) & ", " &
                                    "co_country_del = " & gU.convdbNVCData(gU.dbEncode(CO_COUNTRY_DEL.Text)) & ", " &
                                    "co_cus_cont = " & gU.convdbNVCData(gU.dbEncode(CO_CUS_CONT.Text)) & ", " &
                                    "co_cus_cont_tel = " & gU.convdbNVCData(gU.dbEncode(CO_CUS_CONT_TEL.Text)) & ", " &
                                    "co_rem = " & gU.convdbNVCData(gU.dbEncode(CO_REM.Text)) & ", " &
                                    "co_ftrack_no = " & gU.convdbNVCData(gU.dbEncode(CO_FTRACK_NO.Text)) & ", " &
                                    "co_inv_no = " & gU.convdbNVCData(gU.dbEncode(CO_INV_NO.Text)) & "," &
                                    "route_id = " & gU.convdbNVCData(gU.dbEncode(ROUTE_ID.Text)) & "," &
                                    "co_cust_inv_no = " & gU.convdbNVCData(gU.dbEncode(CO_CUST_INV_NO.Text)) & "," &
                                    "co_ship_mode=" & gU.convdbNVCData(gU.dbEncode(CO_SHIP_MODE.SelectedValue)) & "," &
                                    "CO_TARGET_DELDATE=" & gU.convdbDate(gU.dbEncode(CO_TARGET_DELDATE.Text)) & "," &
                                    "CO_EDI_SIR_NO=" & gU.convdbNVCData(gU.dbEncode(CO_EDI_SIR_NO.Text.Trim)) & "," &
                                    "CO_PROVINCE=" & gU.convdbNVCData(gU.dbEncode(CO_PROVINCE.Text.Trim)) & "," &
                                    "CO_CITY=" & gU.convdbNVCData(gU.dbEncode(CO_CITY.Text.Trim)) & "," &
                                    "CO_SENDER=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER.Text.Trim)) & "," &
                                    "CO_SENDER_COUNTRY=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER_COUNTRY.Text.Trim)) & "," &
                                    "CO_SENDER_PROVINCE=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER_PROVINCE.Text.Trim)) & "," &
                                    "CO_SENDER_REGION=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER_REGION.Text.Trim)) & "," &
                                    "CO_SENDER_ADDR=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER_ADDR.Text.Trim)) & "," &
                                    "CO_SENDER_TEL=" & gU.convdbNVCData(gU.dbEncode(CO_SENDER_TEL.Text.Trim)) & "," &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_code = '" & gU.dbEncode(nextNo) & "' "

                    REM **********************
                    '"co_track_no = " & gU.convdbNVCData(gU.dbEncode(CO_TRACK_NO.Text)) & ", " & _

                    Dim tempHOLDDT As New DataTable

                    tempHOLDDT = gU.getSessionTempData(ModuleAbb, "holdDT", Nothing)
                    Dim updateHOLDSQL As String = ""
                    Dim tempSQL As String = ""

                    Dim maxSeq As String = ""
                    Dim IN_STOCK_QTY As Integer = 0

                    If Not tempHOLDDT Is Nothing AndAlso tempHOLDDT.Rows.Count > 0 Then
                        For i = 0 To tempHOLDDT.Rows.Count - 1
                            updateHOLDSQL = ""
                            tempSQL = ""

                            Select Case tempHOLDDT.Rows(i).Item("mFlag").ToString.Trim
                                Case "N"
                                    If tempHOLDDT.Rows(i).Item("COH_TYPE").ToString.Trim = "STOCK" Then
                                        IN_STOCK_QTY = CInt(gU.decodeEmptyCInt(tempHOLDDT.Rows(i).Item("COH_QTY").ToString.Trim, 0))
                                    Else
                                        IN_STOCK_QTY = 0
                                    End If

                                    maxSeq = DB.getValueFromSQL("Select ISNULL(Max(convert(int,COH_SEQ)) + 1,1) as value from WMS_CUST_ORDER_HOLD " &
                                                                " WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "' " &
                                                                " AND CO_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("CO_CODE").ToString.Trim) & "' AND COD_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "'", gConn, transaction)

                                    updateHOLDSQL = "insert into wms_cust_order_hold (IMP_CODE, STORER_CODE, " &
                                                        "CO_CODE, COD_SEQ, COH_SEQ, " &
                                                        "COH_TYPE, COH_STATUS, RO_CODE, " &
                                                        "ROD_SEQ, COH_QTY, COH_TOT_PCS, " &
                                                        "SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " &
                                                        "COH_IN_STOCK_QTY, COH_BATCH_NO) " &
                                                    "values ('" & Session("IMP_CODE") & "',  '" & gU.dbEncode(tempHOLDDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "', " &
                                                        "'" & gU.dbEncode(tempHOLDDT.Rows(i).Item("CO_CODE").ToString.Trim) & "',  '" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "','" & maxSeq & "', " &
                                                        "'" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_TYPE").ToString.Trim) & "', '" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_STATUS").ToString.Trim) & "',  '" & gU.dbEncode(tempHOLDDT.Rows(i).Item("RO_CODE").ToString.Trim) & "', " &
                                                        "'" & gU.dbEncode(tempHOLDDT.Rows(i).Item("ROD_SEQ").ToString.Trim) & "', '" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_QTY").ToString.Trim) & "',  null, " &
                                                        "'" & Session("usr_id") & "', Getdate(),  Getdate(), '" & Session("usr_id") & "', " &
                                                        "'" & gU.dbEncode(IN_STOCK_QTY) & "', " & gU.convdbNVCData(gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_BATCH_NO").ToString.Trim)) & ") "

                                Case "U"

                                    If tempHOLDDT.Rows(i).Item("COH_STATUS").ToString.Trim <> tempHOLDDT.Rows(i).Item("ORI_STATUS").ToString.Trim AndAlso tempHOLDDT.Rows(i).Item("COH_STATUS").ToString.Trim = "RELEASE" Then
                                        tempSQL &= "COH_REL_DATE=Getdate(), "
                                    End If

                                    If tempHOLDDT.Rows(i).Item("COH_TYPE").ToString.Trim = "STOCK" Then
                                        tempSQL &= "COH_IN_STOCK_QTY='" & tempHOLDDT.Rows(i).Item("COH_QTY").ToString.Trim & "', "
                                    End If

                                    updateHOLDSQL = "UPDATE wms_cust_order_hold SET " &
                                                    " COH_STATUS='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("ORI_STATUS").ToString.Trim) & "'," &
                                                    " RO_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("RO_CODE").ToString.Trim) & "'," &
                                                    " ROD_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("ROD_SEQ").ToString.Trim) & "'," &
                                                    " COH_QTY='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_QTY").ToString.Trim) & "'," &
                                                    " COH_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_BATCH_NO").ToString.Trim)) & ", " &
                                                    tempSQL &
                                                    " SYS_LUB='" & Session("usr_id") & "'," &
                                                    " SYS_LUD=Getdate() " &
                                                    " Where IMP_CODE= '" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "'" &
                                                    " AND CO_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("CO_CODE").ToString.Trim) & "'" &
                                                    " AND COD_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "' AND COH_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_SEQ").ToString.Trim) & "' "
                                Case "D"
                                    updateHOLDSQL = "Delete From wms_cust_order_hold Where " &
                                                    " IMP_CODE= '" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "'" &
                                                    " AND CO_CODE='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("CO_CODE").ToString.Trim) & "'" &
                                                    " AND COD_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COD_SEQ").ToString.Trim) & "' AND COH_SEQ='" & gU.dbEncode(tempHOLDDT.Rows(i).Item("COH_SEQ").ToString.Trim) & "' "

                            End Select


                            If updateHOLDSQL <> "" Then gDB.amendData(updateHOLDSQL, gConn, transaction)

                        Next
                    End If

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        'uiFun.reOrderDetails(dt, "cod_disp_seq")
                        reOrderCOD(dt)

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""

                            cod_batch_no = ""
                            cod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("cod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("COD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("COD_EXPIRY_DATE").ToString.Trim, ""))

                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    If cod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(cod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            insertSql = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        "'" & gU.dbEncode(cod_batch_no) & "'," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(insertSql, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into wms_cust_order_d ( " &
                                                    "co_code, imp_code, storer_code, cod_seq, cod_disp_seq, " &
                                                    "cod_job_no, COD_TICKET_NO, cod_cust_code, " &
                                                    "cod_itm_parent, cod_itm_code, cod_pack_key, cod_itm_desc, " &
                                                    "cod_packing, cod_qty, cod_uom, cod_qty2, cod_uom2, cod_pcs_uom, " &
                                                    "cod_totpcs, cod_tot_wgt, cod_tot_cbm, cod_rem, " &
                                                    "cod_pallet_no, cod_vnd_code, cod_ref_seq, cod_batch_no, " &
                                                    "cod_lot_no, cod_carton_no, cod_plant, COD_WH_CODE, COD_REQ_BY, cod_pcs_carton, cod_carton_pallet, " &
                                                    "cod_expiry_date, cod_manu_date, " &
                                                    "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                    "values ( " &
                                                    gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_disp_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_job_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_TICKET_NO").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_cust_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_parent").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_pack_key").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_desc").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_packing").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty2").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom2").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_uom").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_totpcs").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_wgt").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_cbm").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_rem").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(getLastLot(rows.Item("cod_itm_code").ToString.Trim, STORER_CODE.SelectedValue.ToString, rows.Item("cod_cust_code").ToString.Trim), "000"))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_vnd_code").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_ref_seq").ToString.Trim, "null")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(cod_batch_no)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_lot_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_carton_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_plant").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_WH_CODE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_REQ_BY").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_carton").ToString.Trim, "null")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_carton_pallet").ToString.Trim, "null")) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_cust_order_d " &
                                              "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                              "and co_code = '" & gU.dbEncode(CO_CODE.Text.Trim) & "' " &
                                              "and cod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("cod_seq").ToString.Trim, "")) & "' "

                                    updateHOLDSQL = "Delete from wms_cust_order_hold " &
                                                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                    "and co_code = '" & gU.dbEncode(CO_CODE.Text.Trim) & "' " &
                                                    "and cod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("cod_seq").ToString.Trim, "")) & "' "

                                    gDB.amendData(updateHOLDSQL, gConn, transaction)


                                Case Else
                                    If cod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(cod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            insertSql = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        "'" & gU.dbEncode(cod_batch_no) & "'," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(insertSql, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "update wms_cust_order_d set " &
                                                "cod_disp_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_disp_seq").ToString.Trim, ""))) & ", " &
                                                "cod_job_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_job_no").ToString.Trim, ""))) & ", " &
                                                "COD_TICKET_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_TICKET_NO").ToString.Trim, ""))) & ", " &
                                                "cod_cust_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_cust_code").ToString.Trim, ""))) & ", " &
                                                "cod_itm_parent = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_parent").ToString.Trim, ""))) & ", " &
                                                "cod_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_code").ToString.Trim, ""))) & ", " &
                                                "cod_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_pack_key").ToString.Trim, ""))) & ", " &
                                                "cod_itm_desc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_itm_desc").ToString.Trim, ""))) & ", " &
                                                "cod_packing = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_packing").ToString.Trim, ""))) & ", " &
                                                "cod_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty").ToString.Trim, "0")) & ", " &
                                                "cod_uom = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom").ToString.Trim, ""))) & ", " &
                                                "cod_qty2 = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_qty2").ToString.Trim, "0")) & ", " &
                                                "cod_uom2 = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_uom2").ToString.Trim, ""))) & ", " &
                                                "cod_pcs_uom = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_uom").ToString.Trim, "0")) & ", " &
                                                "cod_totpcs = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_totpcs").ToString.Trim, "0")) & ", " &
                                                "cod_tot_wgt = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_wgt").ToString.Trim, "0")) & ", " &
                                                "cod_tot_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_tot_cbm").ToString.Trim, "0")) & ", " &
                                                "cod_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_rem").ToString.Trim, ""))) & ", " &
                                                "cod_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pallet_no").ToString.Trim, "000"))) & ", " &
                                                "cod_vnd_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_vnd_code").ToString.Trim, ""))) & ", " &
                                                "cod_ref_seq = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_ref_seq").ToString.Trim, "null")) & ", " &
                                                "cod_batch_no = " & gU.convdbNVCData(gU.dbEncode(cod_batch_no)) & ", " &
                                                "cod_lot_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_lot_no").ToString.Trim, ""))) & ", " &
                                                "cod_carton_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_carton_no").ToString.Trim, ""))) & ", " &
                                                "cod_plant = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("cod_plant").ToString.Trim, ""))) & ", " &
                                                "COD_WH_CODE = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_WH_CODE").ToString.Trim, ""))) & ", " &
                                                "COD_REQ_BY = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("COD_REQ_BY").ToString.Trim, ""))) & ", " &
                                                "cod_pcs_carton = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_pcs_carton").ToString.Trim, "null")) & ", " &
                                                "cod_carton_pallet = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("cod_carton_pallet").ToString.Trim, "null")) & ", " &
                                                "COD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                "COD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("COD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                                "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and co_code = '" & gU.dbEncode(nextNo.Trim) & "' " &
                                                "and cod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("cod_seq").ToString.Trim, "")) & "'"


                            End Select
                            REM **********************

                            'if session("usr_id") = "OTSADMIN" then response.write(itemSQL & "<br><br>")
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                            'If Session("usr_type") = "T" Or Session("usr_type") = "C" Then
                            '    SendCustCOMail(CO_CODE.Text.Trim, STORER_CODE.SelectedValue)
                            'End If

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                sql_string = "update wms_cust_order_hold " &
                             "set COH_ITM_CODE = d.COD_ITM_CODE, " &
                                 "COH_PACK_KEY = d.COD_PACK_KEY, " &
                                 "COH_PALLET_NO = d.COD_PALLET_NO " &
                             "from wms_cust_order_d d " &
                             "where wms_cust_order_hold.imp_code = d.imp_code " &
                             "and wms_cust_order_hold.STORER_CODE = d.STORER_CODE " &
                             "and wms_cust_order_hold.CO_CODE = d.CO_CODE " &
                             "and wms_cust_order_hold.COD_SEQ = d.COD_SEQ " &
                             "and wms_cust_order_hold.imp_code = '" & gU.dbEncode(imp_code) & "' " &
                             "and wms_cust_order_hold.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                             "and wms_cust_order_hold.co_code = '" & gU.dbEncode(CO_CODE.Text.Trim) & "' "

                gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    CO_CODE.Text = nextNo
                    ViewState("CO_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    CO_CODE.ForeColor = Drawing.Color.Black
                    CO_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                Call BindGV()
                changeLabel()

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

    Function getLastLot(ByVal ITEM_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String) As String

        Dim LAST_LOT As String = ""
        Dim SQLString As String = "Select iSnuLL(DEST_LOT_NUMBER,'')LAST_LOT from WMS_EBS_TRANS_ITX_SO so Where so.ITEM_ID = '" + ITEM_CODE + "' and so.ACCOUNT_NUMBER='" + CUS_CODE + "' and so.IO_ID = '" + STORER_CODE + "'"
        Dim dtItem As DataTable = gDB.getDataTable(SQLString)
        If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
            LAST_LOT = dtItem.Rows(0)("LAST_LOT")
        End If
        Return LAST_LOT
    End Function
    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim pk_code, storerCode As String
        Dim HoldDt As New DataTable
        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("CO_CODE") <> "" Then
            pk_code = ViewState("CO_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("co_code"))
            storerCode = Server.UrlDecode(Request("storer_code"))
        End If
        REM **********************

        CO_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            CO_CODE.ForeColor = Drawing.Color.Red
            CO_STATUS.Text = "NEW"
            btnAttach.Visible = False
            btnPrtLbl.Visible = False
            btnTrack.Visible = False
            btnPrintPalletLabel.Visible = False
            btnNOTE.Visible = False

            REM **********************

            uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, (CUS_CODE+'-'+CUS_NAME) as CUS_NAME from wms_customer where " &
                                          "storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                        "and imp_code = '" & Session("IMP_CODE") & "' " &
                                        "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

            'uiFun.load_dropdown(CO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " &
            '                              "storer_code = '" & STORER_CODE.SelectedValue & "' " &
            '                              "and imp_code = '" & Session("IMP_CODE") & "' " &
            '                              "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
        Else
            REM **********************
            REM Modify Here
            btnAttach.Visible = True
            btnPrtLbl.Visible = True
            btnTrack.Visible = True
            btnPrintPalletLabel.Visible = True
            btnNOTE.Visible = True

            REM Generate Data Table from Header
            SQLString = " SELECT wms_cust_order.IMP_CODE,  wms_cust_order.STORER_CODE,  wms_cust_order.CO_CODE, " &
                        " wms_cust_order.CO_CUS_REF_NO,  wms_cust_order.CO_PROJECT_NO,  wms_cust_order.CO_STATUS, " &
                        " Convert(varchar, wms_cust_order.CO_DATE," & DDFORMAT & ") as CO_DATE,  wms_cust_order.CUS_CODE,  wms_cust_order.CUS_NAME, " &
                        " wms_cust_order.CO_ADDR1,  wms_cust_order.CO_ADDR2,  wms_cust_order.CO_ADDR3, " &
                        " wms_cust_order.CO_AREA_DEL,  wms_cust_order.CO_DELIVER_TO, wms_cust_order.CO_REGION_DEL,  wms_cust_order.CO_COUNTRY_DEL, " &
                        " wms_cust_order.CO_CUS_CONT,  wms_cust_order.CO_CUS_CONT_TEL,  wms_cust_order.CO_REM, " &
                        " wms_cust_order.CO_FTRACK_NO,  wms_cust_order.CO_RO_CODE, " &
                        " wms_cust_order.CO_TEAM,  wms_cust_order.CO_SCH_DATE,  wms_cust_order.CO_CONTRACT_TYPE, " &
                        " wms_cust_order.CO_AGREEMENT_NO,  wms_cust_order.CO_INSTALLATION_REQ,  wms_cust_order.CO_INITIAL_COLLECT, wms_cust_order.CO_EDI_SIR_NO," &
                        " wms_cust_order.CO_INITIAL_COLLECT_AMT,  wms_cust_order.CO_PAY_TERMS,  wms_cust_order.CO_DELIVERED_BY, wms_cust_order.co_ship_mode,Convert(varchar, wms_cust_order.CO_TARGET_DELDATE," & DDFORMAT & ") as CO_TARGET_DELDATE," &
                        " wms_cust_order.SYS_LUB,  wms_cust_order.SYS_LUD,  wms_cust_order.SYS_CD,  wms_cust_order.SYS_CB, " &
                        " wms_cust_order.CO_INV_NO, wms_cust_order.ROUTE_ID,  wms_cust_order.CO_CUST_INV_NO, wms_storer.STO_BATCH_FIELD_REF, " &
                        " wms_cust_order.CO_DATA_FR, wms_cust_order.CO_PROVINCE, wms_cust_order.CO_CITY, " &
                        " wms_cust_order.CO_SENDER, wms_cust_order.CO_SENDER_COUNTRY, wms_cust_order.CO_SENDER_PROVINCE, wms_cust_order.CO_SENDER_REGION, wms_cust_order.CO_SENDER_ADDR, wms_cust_order.CO_SENDER_TEL " &
                        " FROM wms_cust_order left outer join wms_storer " &
                        "on wms_storer.storer_code = wms_cust_order.storer_code " &
                        "AND wms_storer.imp_code = wms_cust_order.imp_code " &
                        "where wms_cust_order.co_code = '" & gU.dbEncode(pk_code) & "' " &
                        "and wms_cust_order.imp_code = '" & imp_code & "' " &
                        "and wms_cust_order.storer_code = '" & gU.dbEncode(storerCode) & "' "
            'wms_cust_order.CO_TRACK_NO, 

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                'imp_code = dt.Rows(0).Item("IMP_CODE").ToString
                ViewState("STO_BATCH_FIELD_REF") = dt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString

                CO_CODE.Text = dt.Rows(0).Item("CO_CODE").ToString
                CO_STATUS.Text = dt.Rows(0).Item("CO_STATUS").ToString
                CO_CUS_REF_NO.Text = dt.Rows(0).Item("CO_CUS_REF_NO").ToString
                'CO_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("CO_DATE").ToString)
                CO_DATE.Text = dt.Rows(0).Item("CO_DATE").ToString

                CO_DATA_FR.Value = dt.Rows(0).Item("CO_DATA_FR").ToString.Trim

                CO_ADDR1.Text = dt.Rows(0).Item("CO_ADDR1").ToString
                CO_ADDR2.Text = dt.Rows(0).Item("CO_ADDR2").ToString
                CO_ADDR3.Text = dt.Rows(0).Item("CO_ADDR3").ToString
                CO_AREA_DEL.Text = dt.Rows(0).Item("CO_AREA_DEL").ToString
                CO_DELIVER_TO.Text = dt.Rows(0).Item("CO_DELIVER_TO").ToString
                CO_REGION_DEL.Text = dt.Rows(0).Item("CO_REGION_DEL").ToString
                CO_COUNTRY_DEL.Text = dt.Rows(0).Item("CO_COUNTRY_DEL").ToString
                CO_CUS_CONT.Text = dt.Rows(0).Item("CO_CUS_CONT").ToString
                CO_CUS_CONT_TEL.Text = dt.Rows(0).Item("CO_CUS_CONT_TEL").ToString
                CO_REM.Text = dt.Rows(0).Item("CO_REM").ToString
                'CO_TRACK_NO.Text = dt.Rows(0).Item("CO_TRACK_NO").ToString
                CO_FTRACK_NO.Text = dt.Rows(0).Item("CO_FTRACK_NO").ToString
                CO_INV_NO.Text = dt.Rows(0).Item("CO_INV_NO").ToString
                ROUTE_ID.Text = dt.Rows(0).Item("ROUTE_ID").ToString
                CO_CUST_INV_NO.Text = dt.Rows(0).Item("CO_CUST_INV_NO").ToString
                CO_SHIP_MODE.SelectedValue = dt.Rows(0).Item("CO_SHIP_MODE").ToString
                CO_TARGET_DELDATE.Text = dt.Rows(0).Item("CO_TARGET_DELDATE").ToString
                CO_EDI_SIR_NO.Text = dt.Rows(0).Item("CO_EDI_SIR_NO").ToString

                CO_PROVINCE.Text = dt.Rows(0).Item("CO_PROVINCE").ToString
                CO_CITY.Text = dt.Rows(0).Item("CO_CITY").ToString

                CO_SENDER.Text = dt.Rows(0).Item("CO_SENDER").ToString
                CO_SENDER_COUNTRY.Text = dt.Rows(0).Item("CO_SENDER_COUNTRY").ToString
                CO_SENDER_PROVINCE.Text = dt.Rows(0).Item("CO_SENDER_PROVINCE").ToString
                CO_SENDER_REGION.Text = dt.Rows(0).Item("CO_SENDER_REGION").ToString
                CO_SENDER_ADDR.Text = dt.Rows(0).Item("CO_SENDER_ADDR").ToString
                CO_SENDER_TEL.Text = dt.Rows(0).Item("CO_SENDER_TEL").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                'If dt.Rows(0).Item("CO_PROJECT_NO").ToString = "" Then
                '    uiFun.load_dropdown(CO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT  where " &
                '                          " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                '                        "and imp_code = '" & Session("IMP_CODE") & "' " &
                '                        "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
                'Else
                '    uiFun.load_dropdown(CO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT WHERE " &
                '                          " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                '                          "and imp_code = '" & Session("IMP_CODE") & "' " &
                '                          "and PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CO_PROJECT_NO").ToString) & "' ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , , , True)
                'End If
                CO_PROJECT_NO.Text = cU.chgToFullDF(dt.Rows(0).Item("CO_PROJECT_NO").ToString)
                PRJ_NAME.Text = DB.getValueFromSQL("SELECT PRJ_NAME FROM WMS_PROJECT WHERE storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CO_PROJECT_NO").ToString) & "' ")
                'CUS_NAME.Text = DB.getValueFromSQL("SELECT CUS_NAME FROM WMS_CUSTOMER WHERE storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and CUS_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CUS_CODE").ToString) & "' ")
                CUS_NAME.Text = dt.Rows(0).Item("CUS_NAME").ToString
                uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, (CUS_CODE+'-'+CUS_NAME) as CUS_NAME from wms_customer where " &
                                              "storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                              "and imp_code = '" & Session("IMP_CODE") & "' " &
                                              "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

                CUS_CODE.SelectedValue = dt.Rows(0).Item("CUS_CODE").ToString
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        SQLString = " SELECT 'U' as mFlag, wms_cust_order_d.IMP_CODE, wms_cust_order_d.STORER_CODE, wms_cust_order_d.CO_CODE, wms_cust_order_d.COD_SEQ, " &
                        " wms_cust_order_d.COD_JOB_NO, wms_cust_order_d.COD_TICKET_NO, wms_cust_order_d.COD_DISP_SEQ, wms_cust_order_d.COD_CUST_CODE, " &
                        " wms_cust_order_d.COD_CUST_REM, wms_cust_order_d.COD_ITM_PARENT, wms_cust_order_d.COD_ITM_CODE, wms_cust_order_d.COD_PACK_KEY, " &
                        " wms_cust_order_d.COD_PALLET_NO, wms_cust_order_d.COD_ITM_DESC, wms_cust_order_d.cod_post_qty, wms_cust_order_d.COD_PACKING, wms_cust_order_d.COD_QTY, " &
                        " wms_cust_order_d.COD_UOM, wms_cust_order_d.COD_QTY2, wms_cust_order_d.COD_UOM2, wms_cust_order_d.COD_PCS_UOM, wms_cust_order_d.COD_TOTPCS, wms_cust_order_d.COD_DELIV_QTY, " &
                        " wms_cust_order_d.COD_TOT_WGT, wms_cust_order_d.COD_TOT_CBM, wms_cust_order_d.COD_REM, wms_cust_order_d.SYS_LUB, " &
                        " wms_cust_order_d.SYS_LUD, wms_cust_order_d.SYS_CD, wms_cust_order_d.SYS_CB, wms_cust_order_d.COD_VND_CODE, " &
                        " wms_cust_order_d.COD_BATCH_NO, wms_cust_order_d.COD_BATCH_NO as ori_batch, wms_cust_order_d.COD_BATCH_CD, wms_cust_order_d.COD_LOT_NO, wms_cust_order_d.COD_CARTON_NO, " &
                        " wms_cust_order_d.COD_PLANT, wms_cust_order_d.COD_WH_CODE, wms_cust_order_d.COD_REQ_BY, wms_cust_order_d.COD_PCS_CARTON, wms_cust_order_d.COD_CARTON_PALLET, wms_cust_order_d.COD_REF_SEQ, " &
                        " ISNULL(SUM(COH_QTY),0) as HOLD_QTY, wms_cust_order_d.COD_ITM_CODE as h_COD_ITM_CODE, wms_cust_order_d.COD_PACK_KEY as h_COD_PACK_KEY," &
                        " wms_item.ITM_SKU_NO, wms_item.ITM_DESC, " &
                        " Convert(varchar, wms_cust_order_d.COD_EXPIRY_DATE, " & DDFORMAT & ") as COD_EXPIRY_DATE, " &
                        " Convert(varchar, wms_cust_order_d.COD_MANU_DATE, " & DDFORMAT & ") as COD_MANU_DATE " &
                    " FROM wms_cust_order_d left outer join wms_cust_order_hold " &
                    " ON wms_cust_order_d.IMP_CODE  = wms_cust_order_hold.IMP_CODE " &
                    " AND wms_cust_order_d.STORER_CODE = wms_cust_order_hold.STORER_CODE " &
                    " AND wms_cust_order_d.CO_CODE     = wms_cust_order_hold.CO_CODE " &
                    " AND wms_cust_order_d.COD_SEQ     = wms_cust_order_hold.COD_SEQ " &
                    " and wms_cust_order_hold.coh_status = 'HOLD'" &
                    " left outer join wms_item " &
                    " on wms_cust_order_d.IMP_CODE = wms_item.IMP_CODE " &
                    " and wms_cust_order_d.STORER_CODE = wms_item.STORER_CODE " &
                    " and wms_cust_order_d.COD_ITM_CODE = wms_item.ITM_CODE " &
                    " and wms_cust_order_d.COD_PACK_KEY = wms_item.PACK_KEY " &
                    " WHERE wms_cust_order_d.co_code = '" & gU.dbEncode(pk_code) & "' " &
                    " and wms_cust_order_d.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    " and wms_cust_order_d.imp_code = '" & imp_code & "'" &
                    " Group by wms_cust_order_d.IMP_CODE, wms_cust_order_d.STORER_CODE, wms_cust_order_d.CO_CODE, wms_cust_order_d.COD_SEQ, " &
                        " wms_cust_order_d.COD_JOB_NO, wms_cust_order_d.COD_TICKET_NO, wms_cust_order_d.COD_DISP_SEQ, wms_cust_order_d.COD_CUST_CODE, " &
                        " wms_cust_order_d.COD_CUST_REM, wms_cust_order_d.COD_ITM_PARENT, wms_cust_order_d.COD_ITM_CODE, wms_cust_order_d.COD_PACK_KEY, " &
                        " wms_cust_order_d.COD_PALLET_NO, wms_cust_order_d.COD_ITM_DESC, wms_cust_order_d.cod_post_qty, wms_cust_order_d.COD_PACKING, wms_cust_order_d.COD_QTY, " &
                        " wms_cust_order_d.COD_UOM, wms_cust_order_d.COD_QTY2, wms_cust_order_d.COD_UOM2, wms_cust_order_d.COD_PCS_UOM, wms_cust_order_d.COD_TOTPCS, wms_cust_order_d.COD_DELIV_QTY, " &
                        " wms_cust_order_d.COD_TOT_WGT, wms_cust_order_d.COD_TOT_CBM, wms_cust_order_d.COD_REM, wms_cust_order_d.SYS_LUB, " &
                        " wms_cust_order_d.SYS_LUD, wms_cust_order_d.SYS_CD, wms_cust_order_d.SYS_CB, wms_cust_order_d.COD_VND_CODE, " &
                        " wms_cust_order_d.COD_BATCH_NO, wms_cust_order_d.COD_BATCH_CD, wms_cust_order_d.COD_LOT_NO, wms_cust_order_d.COD_CARTON_NO, " &
                        " wms_cust_order_d.COD_PLANT, wms_cust_order_d.COD_WH_CODE, wms_cust_order_d.COD_REQ_BY, wms_cust_order_d.COD_PCS_CARTON, wms_cust_order_d.COD_CARTON_PALLET, wms_cust_order_d.COD_REF_SEQ, wms_item.ITM_SKU_NO, wms_item.ITM_DESC, " &
                        " wms_cust_order_d.cod_expiry_date, wms_cust_order_d.cod_manu_date "


        SQLString = SQLString & " order by wms_cust_order_d.cod_disp_seq, Convert(int,wms_cust_order_d.cod_seq) "
        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        'ViewState("dt") = dt
        gU.setSessionTempData(ModuleAbb, "codDT", dt)
        GridView1.DataBind()



        SQLString = " SELECT 'U' as mFlag, 0 as rod_qty, wms_cust_order_hold.IMP_CODE, wms_cust_order_hold.STORER_CODE, wms_cust_order_hold.CO_CODE, " &
                    "  wms_cust_order_hold.COD_SEQ, wms_cust_order_hold.COH_SEQ, wms_cust_order_hold.COH_TYPE, wms_cust_order_hold.COH_BATCH_NO, " &
                    "  wms_cust_order_hold.COH_STATUS, wms_cust_order_hold.COH_STATUS as ori_status, wms_cust_order_hold.RO_CODE, wms_cust_order_hold.ROD_SEQ, wms_cust_order_hold.COH_IN_STOCK_QTY, wms_cust_order_hold.COH_REL_QTY, " &
                    "  wms_cust_order_hold.COH_QTY, wms_cust_order_hold.COH_TOT_PCS, " &
                    " Convert(varchar, wms_cust_order_hold.COH_REL_DATE," & DDFORMAT & ") as COH_REL_DATE, " &
                    " wms_cust_order_hold.COH_LOC, " &
                    " Convert(varchar, wms_cust_order_hold.sys_cd," & DDFORMAT & ") as sys_cd" &
                    " FROM wms_cust_order_hold " &
                    " where wms_cust_order_hold.co_code = '" & gU.dbEncode(pk_code) & "' " &
                    " and wms_cust_order_hold.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    " and wms_cust_order_hold.imp_code = '" & imp_code & "'"

        HoldDt = gDB.getDataTable(SQLString)

        gU.setSessionTempData(ModuleAbb, "holdDT", HoldDt)

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        CO_STATUS.Text = "CANCELLED"
        Call save()
        ar.sec_write = "N"
        CancelBtn.Visible = False
        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub reOpenBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles reOpenBtn.Click
        Dim updtSql, SQLString As String

        Dim gConn As SqlConnection

        Dim dt As DataTable

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try

            updtSql = "update WMS_CUST_ORDER " &
                        "set CO_STATUS = 'NEW', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and CO_CODE = '" & gU.dbEncode(CO_CODE.Text) & "' " &
                        "AND NOT EXISTS (SELECT 1 FROM WMS_CUST_ORDER_D D WHERE ISNULL(COD_POST_QTY,0) <> 0 " &
                        "AND D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE) "

            gDB.amendData(updtSql, gConn, transaction)

            updtSql = "update WMS_CUST_ORDER " &
                        "set CO_STATUS = 'PARTIAL', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and CO_CODE = '" & gU.dbEncode(CO_CODE.Text) & "' " &
                        "AND EXISTS (SELECT 1 FROM WMS_CUST_ORDER_D D WHERE ISNULL(COD_POST_QTY,0) <> 0 " &
                        "AND D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE) "

            gDB.amendData(updtSql, gConn, transaction)

            transaction.Commit()

            SQLString = "select CO_STATUS FROM WMS_CUST_ORDER M " &
                                    "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(CO_CODE.Text) & "' "
            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                CO_STATUS.Text = dt.Rows(0).Item(0).ToString.Trim
            End If

            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "", "CO has been re-opened.", Session("gLang"))

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

    Protected Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Dim cancelSql As String = "update wms_cust_order " &
                     "set co_status = 'CLOSED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(imp_code.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and co_code = '" & gU.dbEncode(CO_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            CO_STATUS.Text = "CLOSED"

            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "", "CO has been closed.", Session("gLang"))

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

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"

        If dt.Rows(e.RowIndex).Item("cod_ref_seq").ToString = "" Then
            Dim splitRows As DataRow() = dt.Select("cod_ref_seq = " & dt.Rows(e.RowIndex).Item("cod_seq").ToString)

            For Each r As DataRow In splitRows
                r.Item("mFlag") = "D"
                r.AcceptChanges()
            Next
            dt.AcceptChanges()

            gU.setSessionTempData(ModuleAbb, "codDT", dt)

            GridView1.DataSource = dt
            GridView1.DataBind()
        Else
            dt.AcceptChanges()
        End If
    End Sub

    'Protected Sub CO_PROJECT_NO_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CO_PROJECT_NO.SelectedIndexChanged
    '    PRJ_NAME.Text = DB.getValueFromSQL("select prj_name from wms_project where prj_code = '" & CO_PROJECT_NO.SelectedValue & "' and imp_code = '" & Session("IMP_CODE") & "'")
    'End Sub

    Protected Sub CUS_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_CODE.SelectedIndexChanged
        Dim SQLSTR As String = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " &
                                "cus_area_del, cus_region_del, cus_country_del, cus_cont_per_ord, cus_cont_tel_ord " &
                                "from wms_customer where cus_code = '" & gU.dbEncode(CUS_CODE.SelectedValue) & "' " &
                                "and storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                "and imp_code = '" & Session("IMP_CODE") & "'"
        Dim cust_table As New DataTable

        cust_table = gDB.getDataTable(SQLSTR)

        If cust_table.Rows.Count > 0 Then
            CUS_NAME.Text = cust_table.Rows(0).Item("cus_name").ToString
            CO_ADDR1.Text = cust_table.Rows(0).Item("cus_addr1_del").ToString
            CO_ADDR2.Text = cust_table.Rows(0).Item("cus_addr2_del").ToString
            CO_ADDR3.Text = cust_table.Rows(0).Item("cus_addr3_del").ToString
            CO_AREA_DEL.Text = cust_table.Rows(0).Item("cus_area_del").ToString
            CO_REGION_DEL.Text = cust_table.Rows(0).Item("cus_region_del").ToString
            CO_COUNTRY_DEL.Text = cust_table.Rows(0).Item("cus_country_del").ToString()
            CO_CUS_CONT.Text = cust_table.Rows(0).Item("cus_cont_per_ord").ToString()
            CO_CUS_CONT_TEL.Text = cust_table.Rows(0).Item("cus_cont_tel_ord").ToString()
        Else
            CUS_NAME.Text = ""
            CO_ADDR1.Text = ""
            CO_ADDR2.Text = ""
            CO_ADDR3.Text = ""
            CO_AREA_DEL.Text = ""
            CO_REGION_DEL.Text = ""
            CO_COUNTRY_DEL.Text = ""
            CO_CUS_CONT.Text = ""
            CO_CUS_CONT_TEL.Text = ""
        End If
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        PRJ_NAME.Text = ""
        CUS_NAME.Text = ""
        CO_ADDR1.Text = ""
        CO_ADDR2.Text = ""
        CO_ADDR3.Text = ""
        CO_AREA_DEL.Text = ""
        CO_REGION_DEL.Text = ""
        CO_COUNTRY_DEL.Text = ""
        CO_CUS_CONT.Text = ""
        CO_CUS_CONT_TEL.Text = ""

        uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, (CUS_CODE+'-'+CUS_NAME) as CUS_NAME  from wms_customer where " &
                   " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                 "and imp_code = '" & Session("IMP_CODE") & "' " &
                 "and cus_status = 'ACTIVE' " &
                 "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

        'uiFun.load_dropdown(CO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " &
        '                                  " storer_code = '" & STORER_CODE.SelectedValue & "' " &
        '                                "and imp_code = '" & Session("IMP_CODE") & "' " &
        '                                "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))

        setDefStorerInfo()
    End Sub

    Protected Sub setDefStorerInfo()
        ViewState("STORER_CODE") = STORER_CODE.SelectedValue

        Dim SQLString As String = ""
        Dim ldt As New DataTable

        SQLString = "SELECT STO_BATCH_FIELD_REF from WMS_STORER " &
                    "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and IMP_CODE = '" & gU.dbEncode(imp_code) & "'"

        REM **********************
        ldt = gDB.getDataTable(SQLString)
        If ldt.Rows.Count > 0 Then
            ViewState("STO_BATCH_FIELD_REF") = ldt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString.Trim
        End If
        ldt = Nothing
    End Sub

    Protected Sub addItemtoCO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(Convert(int,COD_SEQ)) + 1 from WMS_CUST_ORDER_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and CO_CODE = '" & gU.dbEncode(CO_CODE.Text.Trim) & "' "
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
            Dim CO_dt As DataTable
            Dim warn_msg As String = ""

            Dim itemPackList As String = ""
            Dim locPackList As String = ""

            'Dim itemListarray As String()
            'Dim packKeyListarray As String()
            'Dim vndCodeListarray As String()

            Dim itemKeyListArray As String()
            Dim locKeyListArray As String()
            Dim qtyListArray As String()

            itemKeyListArray = Split(itemKeyList.Value, ", ")
            locKeyListArray = Split(locKeyList.Value, ", ")
            qtyListArray = Split(qtyList.Value, ", ")


            'itemListarray = Split(itemList.Value, ", ")
            'packKeyListarray = Split(packKeyList.Value, ", ")
            'vndCodeListarray = Split(vndList.Value, ", ")

            'If itemListarray.Count = 0 Then
            '    If itemList.Value <> "" Then
            '        itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
            '    End If
            'Else
            '    For i = 0 To itemListarray.Count - 1
            '        itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
            '    Next
            'End If

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i As Integer = 0 To itemKeyListArray.Count - 1
                    If qtyitemDict.ContainsKey(itemKeyListArray(i) & "#_#" & Trim(locKeyListArray(i))) Then
                        qtyitemDict(itemKeyListArray(i) & "#_#" & Trim(locKeyListArray(i))) = CStr(gU.decodeEmptyCInt(qtyitemDict(itemKeyListArray(i) & "#_#" & Trim(locKeyListArray(i))), 0) + gU.decodeEmptyCInt(qtyListArray(i), 0))
                    Else
                        qtyitemDict.Add(itemKeyListArray(i) & "#_#" & Trim(locKeyListArray(i)), qtyListArray(i))
                    End If
                Next
            End If

            If itemKeyListArray.Count = 0 Then
                If itemKeyList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemKeyList.Value)
                    locPackList = Server.HtmlDecode(itemKeyList.Value) & "#_#" & Server.HtmlDecode(Trim(locKeyList.Value))
                End If
            Else
                For i = 0 To itemKeyListArray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemKeyListArray(i)))
                    locPackList = gU.appendToList(locPackList, Server.HtmlDecode(itemKeyListArray(i)) & "#_#" & Server.HtmlDecode(Trim(locKeyListArray(i))))
                Next
            End If

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

            'SQLString = SQLString & " order by wms_item.itm_code, wms_item.pack_key, wms_alt_vend_item.vnd_code"


            '"WMS_ALT_VEND_ITEM.AITM_PCS_PER_PACK, WMS_ALT_VEND_ITEM.AITM_VOL, WMS_ALT_VEND_ITEM.AITM_CBM, " & _

            '"LEFT OUTER JOIN WMS_ALT_VEND_ITEM " & _
            '            "ON WMS_ITEM.IMP_CODE = WMS_ALT_VEND_ITEM.IMP_CODE " & _
            '                "and WMS_ITEM.STORER_CODE = WMS_ALT_VEND_ITEM.STORER_CODE " & _
            '                "and WMS_ITEM.ITM_CODE = WMS_ALT_VEND_ITEM.ITM_CODE " & _
            '                "and WMS_ITEM.PACK_KEY = WMS_ALT_VEND_ITEM.PACK_KEY " & _

            '"AND ISNULL(WMS_ALT_VEND_ITEM.VND_CODE, '') = WMS_ITEM_LOC_BAL.VND_CODE " & _

            SQLString = "select WMS_ITEM.ITM_PARENT, WMS_ITEM.ITM_CODE, WMS_ITEM.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_PCS_PER_UOM, " &
                            "WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_DESC, v.AITM_QTY_PER_CTN, v.AITM_VOL, v.CARTON_CBM, " &
                            "WMS_ITEM.ITM_UOM2, " &
                            "WMS_ITEM_LOC_BAL.ITM_CODE + '#_#' + " &
                                "WMS_ITEM_LOC_BAL.PACK_KEY + '#_#' + " &
                                "isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '') + '#_#' + " &
                                "isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') as ITM_LOC_KEY, " &
                                "Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, " & gU.dbEncode(DDFORMAT) & ") as ILOC_EXPIRY_DATE, " &
                                "Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, " & gU.dbEncode(DDFORMAT) & ") as ILOC_MANU_DATE " &
                        "FROM WMS_ITEM " &
                        "LEFT OUTER JOIN V_ALT_VEND_ITEM v " &
                        "on WMS_ITEM.IMP_CODE = v.IMP_CODE " &
                            "and WMS_ITEM.STORER_CODE = v.STORER_CODE " &
                            "and WMS_ITEM.ITM_CODE = v.ITM_CODE " &
                            "and WMS_ITEM.PACK_KEY = v.PACK_KEY " &
                        "LEFT OUTER JOIN ( " &
                            "select imp_code, " &
                                "storer_code, " &
                                "ITM_CODE, " &
                                "PACK_KEY, ILOC_EXPIRY_DATE, ILOC_MANU_DATE," &
                                "ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                                "ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, " &
                                "ISNULL(VND_CODE, '') AS VND_CODE, " &
                                "sum(ILOC_BAL_QTY) as TOTAL_BAL_QTY " &
                            "from WMS_ITEM_LOC_BAL " &
                            "where ISNULL(ILOC_BAL_QTY, 0) > 0 " &
                            "and Ltrim(RTRIM(WMS_ITEM_LOC_BAL.ITM_CODE + '#_#' + " &
                                "WMS_ITEM_LOC_BAL.PACK_KEY + '#_#' + " &
                                "isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '') + '#_#' + " &
                                "isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, ''))) in ('" & Replace(locPackList, ", ", "', '") & "') " &
                            "group by imp_code, storer_code, ITM_CODE, PACK_KEY,ILOC_EXPIRY_DATE, ILOC_MANU_DATE, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, ''), ISNULL(VND_CODE, '') " &
                            ") WMS_ITEM_LOC_BAL " &
                        "on WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE " &
                            "AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " &
                            "AND WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE " &
                            "AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " &
                        "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and wms_item.imp_code = '" & Session("IMP_CODE") & "' " &
                        "and wms_item.itm_code + '#_#' + wms_item.pack_key in ('" & Replace(itemPackList, ", ", "', '") & "') " &
                        "order by WMS_ITEM.ITM_CODE, WMS_ITEM.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO "

            'Response.Write(SQLString)
            REM **********************
            CO_dt = gDB.getDataTable(SQLString)
            For i As Integer = 0 To CO_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                Dim itmQty As Integer = 0

                If qtyitemDict.ContainsKey(CO_dt.Rows(i).Item("ITM_LOC_KEY").ToString.Trim) Then
                    'Each item should only add one row
                    If qtyitemDict(CO_dt.Rows(i).Item("ITM_LOC_KEY").ToString.Trim) = "-999" Then
                        Continue For
                    Else
                        itmQty = gU.decodeEmptyCInt(qtyitemDict(CO_dt.Rows(i).Item("ITM_LOC_KEY").ToString.Trim), 0)
                        qtyitemDict(CO_dt.Rows(i).Item("ITM_LOC_KEY").ToString.Trim) = "-999"
                    End If
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("COD_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("COD_DISP_SEQ") = ViewState("n_cur_seq").ToString

                dt.Rows(rows_count - 1).Item("COD_ITM_PARENT") = CO_dt.Rows(i).Item("ITM_PARENT").ToString
                dt.Rows(rows_count - 1).Item("COD_ITM_CODE") = CO_dt.Rows(i).Item("ITM_CODE").ToString

                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CO_dt.Rows(i).Item("ITM_SKU_NO").ToString

                dt.Rows(rows_count - 1).Item("COD_PACK_KEY") = CO_dt.Rows(i).Item("PACK_KEY").ToString
                dt.Rows(rows_count - 1).Item("COD_ITM_DESC") = CO_dt.Rows(i).Item("ITM_NAME").ToString

                dt.Rows(rows_count - 1).Item("ITM_DESC") = CO_dt.Rows(i).Item("ITM_DESC").ToString

                'dt.Rows(rows_count - 1).Item("COD_QTY") = gU.decodeEmptyCdbl(CO_dt.Rows(i).Item("ITM_BALANCE").ToString, 0)
                dt.Rows(rows_count - 1).Item("COD_QTY") = itmQty


                dt.Rows(rows_count - 1).Item("COD_UOM") = CO_dt.Rows(i).Item("ITM_UOM").ToString

                dt.Rows(rows_count - 1).Item("COD_UOM2") = CO_dt.Rows(i).Item("ITM_UOM2").ToString


                dt.Rows(rows_count - 1).Item("COD_PALLET_NO") = CO_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("COD_BATCH_NO") = CO_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim

                dt.Rows(rows_count - 1).Item("COD_EXPIRY_DATE") = CO_dt.Rows(i).Item("ILOC_EXPIRY_DATE").ToString.Trim
                dt.Rows(rows_count - 1).Item("COD_MANU_DATE") = CO_dt.Rows(i).Item("ILOC_MANU_DATE").ToString.Trim

                If CO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString.Trim <> "" Then
                    dt.Rows(rows_count - 1).Item("COD_PCS_UOM") = gU.decodeEmptyCInt(CO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, 0)
                    dt.Rows(rows_count - 1).Item("COD_TOTPCS") = itmQty * CDbl(gU.decodeEmptyCInt(CO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, 0))
                ElseIf CO_dt.Rows(i).Item("ITM_UOM").ToString = "PCS" Then
                    dt.Rows(rows_count - 1).Item("COD_PCS_UOM") = 1
                    dt.Rows(rows_count - 1).Item("COD_TOTPCS") = itmQty
                End If

                dt.Rows(rows_count - 1).Item("COD_PCS_CARTON") = gU.decodeEmptyCdbl(CO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString, 0)

                'dt.Rows(rows_count - 1).Item("COD_PCS_UOM") = gU.decodeEmptyCInt(CO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, 0)

                'dt.Rows(rows_count - 1).Item("COD_TOTPCS") = gU.decodeEmptyCInt(CO_dt.Rows(i).Item("ITM_BALANCE").ToString, 0) * gU.decodeEmptyCInt(CO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, 0)
                'dt.Rows(rows_count - 1).Item("COD_TOTPCS") = itmQty * gU.decodeEmptyCInt(CO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, 0)

                dt.Rows(rows_count - 1).Item("COD_TOT_WGT") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("AITM_VOL"), "0"), 0)
                dt.Rows(rows_count - 1).Item("COD_TOT_CBM") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("CARTON_CBM"), "0"), 0)


                'dt.Rows(rows_count - 1).Item("COD_VND_CODE") = CO_dt.Rows(i).Item("VND_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("HOLD_QTY") = 0

                'If gU.decodeEmptyCdbl(CO_dt.Rows(i).Item("ITM_BALANCE").ToString, 0) = 0 Then
                '    If warn_msg = "" Then
                '        warn_msg = CO_dt.Rows(i).Item("ITM_CODE").ToString & "-" & CO_dt.Rows(i).Item("VND_CODE").ToString
                '    Else
                '        warn_msg = warn_msg & ", " & _
                '                    CO_dt.Rows(i).Item("ITM_CODE").ToString & "-" & CO_dt.Rows(i).Item("VND_CODE").ToString
                '    End If
                'End If
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
            Next
            dt.AcceptChanges()
            'ViewState("dt") = dt
            gU.setSessionTempData(ModuleAbb, "codDT", dt)
            GridView1.DataSource = dt
            GridView1.DataBind()

            'If warn_msg <> "" Then
            '    If Session("gLang") = "E" Then
            '        uiFun.displayMsg(Me, "", "Zero Qty Items Found!\n\rItem: " & warn_msg, Session("gLang"))
            '    Else
            '        uiFun.displayMsg(Me, "", "沒有存貨!\n\r物件: " & warn_msg, Session("gLang"))
            '    End If
            'End If
        End If
    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here

        'Call cU.newChangeGVLabel(GridView1, "No.", "編號")
        'Call cU.newChangeGVLabel(GridView1, "Job No.", "工作編號")
        Call cU.newChangeGVLabel(GridView1, "Seq.", "序號")
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "Subinventory", "子庫存")
        Call cU.newChangeGVLabel(GridView1, "Req By", "要求人")
        Call cU.newChangeGVLabel(GridView1, "Plan", "計劃")
        Call cU.newChangeGVLabel(GridView1, "Ticket No.", "参考編號")
        Call cU.newChangeGVLabel(GridView1, "Carton No.", "箱號")
        Call cU.newChangeGVLabel(GridView1, "Pallet No.", "板號")
        Call cU.newChangeGVLabel(GridView1, "Lot", "批號")
        Call cU.newChangeGVLabel(GridView1, "Lot No.", "批號.")
        Call cU.newChangeGVLabel(GridView1, "Vendor Code", "供應商代碼")
        Call cU.newChangeGVLabel(GridView1, "Internal Item Code WMS", "貨物編號")
        Call cU.newChangeGVLabel(GridView1, "Item Code", "貨物編號")
        Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        Call cU.newChangeGVLabel(GridView1, "Expiry Date", "失效日期")
        Call cU.newChangeGVLabel(GridView1, "Manufactory Date", "生產日期")
        Call cU.newChangeGVLabel(GridView1, "Pack Type", "封裝形式")
        Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        Call cU.newChangeGVLabel(GridView1, "Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "OS Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "Qty/ Carton", "數量/箱")
        Call cU.newChangeGVLabel(GridView1, "Carton/ Pallet", "數量/板")
        Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        Call cU.newChangeGVLabel(GridView1, "Qty2", "數量2")
        Call cU.newChangeGVLabel(GridView1, "UOM2", "單位2")
        Call cU.newChangeGVLabel(GridView1, "Number/ UOM", "單位件數")
        Call cU.newChangeGVLabel(GridView1, "Total Number", "總件數")
        Call cU.newChangeGVLabel(GridView1, "Weight(kg)", "總重量")
        Call cU.newChangeGVLabel(GridView1, "CBM", "總體積(cbm)")
        Call cU.newChangeGVLabel(GridView1, "Remarks", "備註")
        Call cU.newChangeGVLabel(GridView1, "Hold Qty.", "保持數量")
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Public Sub GetUnitPrice(ByVal Price As Double)
        TotalUnitPrice += Price
    End Sub

    Public Function GetTotal() As Double
        Return TotalUnitPrice
    End Function

    Private Sub reOrderCOD(ByRef impDt As DataTable)
        Dim tempDT As New DataTable

        tempDT = impDt.Clone

        Dim dispSeq As Integer = 1

        Dim rowsFound = From c In impDt.Rows
                        Order By
                        Convert.ToInt32(c.item("cod_disp_seq"))
                        Ascending
                        Select c

        Dim sortRows As Object()

        sortRows = rowsFound.ToArray

        For i As Integer = 0 To sortRows.Count - 1
            If sortRows(i).item("cod_ref_seq").ToString = "" Then
                sortRows(i).Item("cod_disp_seq") = dispSeq

                sortRows(i).AcceptChanges()

                dispSeq += 1

                tempDT.ImportRow(sortRows(i))

                If sortRows(i).Item("mFlag").ToString = "D" Then
                    Dim splitRows As DataRow() = impDt.Select("cod_ref_seq = '" & sortRows(i).item("cod_seq").ToString & "'")

                    For Each r As DataRow In splitRows
                        r.Item("mFlag") = "D"
                        r.AcceptChanges()
                    Next
                End If

            ElseIf sortRows(i).item("cod_ref_seq").ToString <> "" Then
                Dim refRows As DataRow() = impDt.Select("cod_seq = '" & sortRows(i).item("cod_ref_seq").ToString & "'")

                If refRows.Count > 0 Then
                    sortRows(i).Item("cod_disp_seq") = refRows(0).Item("cod_disp_seq")
                    sortRows(i).AcceptChanges()

                    tempDT.ImportRow(sortRows(i))
                End If
            End If
        Next

        tempDT.AcceptChanges()

        impDt = tempDT
    End Sub

    Protected Sub btnExportTF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportTF.Click
        Dim coItemTbl As New DataTable
        Dim coitemSQL As String = ""

        Dim pk_code, storerCode As String

        If ViewState("CO_CODE") <> "" Then
            pk_code = ViewState("CO_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("co_code"))
            storerCode = Server.UrlDecode(Request("storer_code"))
        End If

        coitemSQL += "select distinct cod_plant, COD_TICKET_NO, alv_vnd_itmcode, wms_item.itm_code, cod_vnd_code, "
        coitemSQL += "cod_batch_no, cod_lot_no, cod_qty, cod_carton_no, cod_pallet_no, cod_pcs_carton, cod_carton_pallet, cod_pcs_uom "
        coitemSQL += "from wms_cust_order inner join wms_cust_order_d on "
        coitemSQL += "wms_cust_order.imp_code = wms_cust_order_d.imp_code "
        coitemSQL += "and wms_cust_order.storer_code = wms_cust_order_d.storer_code "
        coitemSQL += "and wms_cust_order.co_code = wms_cust_order_d.co_code "
        coitemSQL += "left outer join wms_item on "
        coitemSQL += "wms_cust_order_d.storer_code = wms_item.storer_code "
        coitemSQL += "and wms_cust_order_d.imp_code = wms_item.imp_code "
        coitemSQL += "and wms_cust_order_d.cod_pack_key = wms_item.pack_key "
        coitemSQL += "and wms_cust_order_d.cod_itm_code = wms_item.itm_code  "
        coitemSQL += "left outer join wms_alt_vend_item on "
        coitemSQL += "wms_item.imp_code = wms_alt_vend_item.imp_code "
        coitemSQL += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
        coitemSQL += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
        coitemSQL += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
        coitemSQL += "and wms_cust_order_d.cod_vnd_code = wms_alt_vend_item.vnd_code "
        coitemSQL += "where wms_cust_order_d.co_code = '" & gU.dbEncode(pk_code) & "' "
        coitemSQL += "and wms_cust_order_d.storer_code = '" & gU.dbEncode(storerCode) & "' "
        coitemSQL += "and wms_cust_order_d.imp_code = '" & imp_code & "' "
        coitemSQL += "and wms_cust_order_d.cod_qty > 0 "
        coitemSQL += "and wms_cust_order_d.cod_pcs_carton > 0 "
        coitemSQL += "and wms_cust_order_d.cod_carton_pallet > 0 "
        coitemSQL += "order by COD_TICKET_NO, cod_plant, wms_item.itm_code, cod_batch_no, cod_lot_no, cod_carton_no, cod_pallet_no "

        coItemTbl = gDB.getDataTable(coitemSQL)

        '****
        'WMS_CUST_ORDER_D.COD_VND_CODE is missing!!!!
        '****

        If coItemTbl.Rows.Count > 0 Then
            Try

                Dim storingPath As String = ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")
                Dim plan As String = ""
                Dim vndCode As String = ""
                Dim fileName As String = ""

                If vndCode = "" Then
                    vndCode = gU.decodeNullOrEmpty(coItemTbl.Rows(0).Item("cod_vnd_code").ToString, "")
                End If

                If plan = "" Then
                    plan = gU.decodeNullOrEmpty(coItemTbl.Rows(0).Item("cod_plant").ToString, "")
                End If

                fileName = vndCode & Format(Now, "ddMMyyyy").ToString & plan

                Dim dirInfo As DirectoryInfo = New IO.DirectoryInfo(storingPath)
                Dim fileList As FileInfo() = dirInfo.GetFiles(fileName & "*.txt")

                Dim seq As String = "01"

                If fileList.Length > 0 Then
                    For x As Integer = 0 To fileList.Length - 1
                        Dim tmpSEQ As String = ""

                        tmpSEQ = fileList(x).Name.Substring(fileList(x).Name.Length - 6, 2)

                        If Int(tmpSEQ) >= CInt(seq) Then
                            seq = (CInt(tmpSEQ) + 1).ToString
                        End If
                    Next

                    For x As Integer = 0 To fileList.Length - 1
                        If fileList(x).CreationTime < Now Then
                            fileList(x).Delete()
                        End If
                    Next
                End If

                If seq.Length = 1 Then seq = "0" & seq

                fileName += seq

                If IO.File.Exists(storingPath & "\" & fileName & ".txt") Then
                    IO.File.Delete(storingPath & "\" & fileName & ".txt")
                End If

                Using fs As New FileStream(storingPath & "\" & fileName & ".txt", FileMode.Create)
                    Using sw As New StreamWriter(fs, Encoding.UTF8)
                        sw.WriteLine("Plan,Reel_ID,Inventec_PN,Mfg_PN,Supplier,DATE_CODE,LOT_NO,Reel_QTY,BOX_ID,Carton_ID")
                        sw.WriteLine("")

                        Dim pc_seq As Integer = 1
                        Dim cp_seq As Integer = 1
                        Dim cartonCnt As Integer = 0
                        Dim itemCnt As Integer = 0

                        Dim lastCartonID As String = ""
                        Dim lastPalletID As String = ""

                        Dim lastpcsCarton As Integer = -1
                        Dim lastCartonPallet As Integer = -1

                        For i As Integer = 0 To coItemTbl.Rows.Count - 1
                            If lastCartonID <> coItemTbl.Rows(i).Item("cod_carton_no").ToString OrElse _
                                lastPalletID <> coItemTbl.Rows(i).Item("cod_pallet_no").ToString Then

                                lastCartonID = gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_carton_no").ToString, "")
                                lastPalletID = gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_pallet_no").ToString, "")

                                pc_seq = 1
                                cp_seq = 1
                                lastpcsCarton = -1
                                lastCartonPallet = -1
                            End If

                            'cartonCnt = 0
                            'itemCnt = 0

                            Dim pcsCarton As Integer = gU.decodeEmptyCInt(coItemTbl.Rows(i).Item("cod_pcs_carton").ToString, 1)
                            Dim CartonPallet As Integer = gU.decodeEmptyCInt(coItemTbl.Rows(i).Item("cod_carton_pallet").ToString, 1)

                            If lastpcsCarton = -1 Then
                                lastpcsCarton = pcsCarton
                            End If

                            If lastCartonPallet = -1 Then
                                lastCartonPallet = CartonPallet
                            End If

                            If pcsCarton > 0 AndAlso CartonPallet > 0 Then
                                Dim init_reel_id As String = gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("COD_TICKET_NO").ToString, "")
                                Dim total_qty As Integer = gU.decodeEmptyCInt(coItemTbl.Rows(i).Item("cod_qty").ToString, 0)
                                Dim tempID As String = ""
                                Dim preFix As String = ""

                                For le As Integer = init_reel_id.Length - 1 To 0 Step -1
                                    If IsNumeric(init_reel_id.Substring(le, 1)) Then
                                        tempID = init_reel_id.Substring(le, 1) & tempID
                                    Else
                                        preFix = init_reel_id.Substring(0, le + 1)
                                        Exit For
                                    End If
                                Next

                                Dim startID As Long = gU.decodeEmptyCLng(tempID, 1)

                                For cnt As Integer = 1 To total_qty
                                    Dim cartonID As String = gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("COD_CARTON_NO").ToString, "")
                                    Dim palletID As String = gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("COD_PALLET_NO").ToString, "")

                                    Dim lineStr As String = ""

                                    If palletID <> "" Then
                                        palletID += "-" & cp_seq.ToString.PadLeft(3, "0")
                                    Else
                                        palletID += cp_seq.ToString.PadLeft(3, "0")
                                    End If

                                    If cartonID <> "" Then
                                        cartonID += "-" & cp_seq.ToString.PadLeft(3, "0") & "-" & pc_seq.ToString.PadLeft(4, "0")
                                    Else
                                        cartonID += cp_seq.ToString.PadLeft(3, "0") & "-" & pc_seq.ToString.PadLeft(4, "0")
                                    End If

                                    lineStr += gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_plant").ToString, "") & ","
                                    lineStr += preFix & startID & ","
                                    lineStr += gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("alv_vnd_itmcode").ToString, "") & ","
                                    lineStr += gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("itm_code").ToString, "") & ","
                                    lineStr += gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_vnd_code").ToString, "") & ","
                                    lineStr += "D" & gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_lot_no").ToString, "") & ","
                                    lineStr += "L" & gU.decodeNullOrEmpty(coItemTbl.Rows(i).Item("cod_batch_no").ToString, "") & ","
                                    lineStr += gU.decodeEmptyCInt(coItemTbl.Rows(i).Item("cod_pcs_uom").ToString, 0) & ","
                                    lineStr += cartonID & ","
                                    lineStr += palletID

                                    sw.WriteLine(lineStr)

                                    startID += 1
                                    itemCnt += 1

                                    If itemCnt <> 0 Then
                                        If itemCnt = lastpcsCarton Then
                                            pc_seq += 1
                                            cartonCnt += 1
                                            itemCnt = 0
                                            lastpcsCarton = -1

                                        ElseIf lastpcsCarton = -1 AndAlso itemCnt Mod pcsCarton = 0 Then
                                            pc_seq += 1
                                            cartonCnt += 1
                                            itemCnt = 0
                                        End If
                                    End If

                                    If cartonCnt <> 0 Then
                                        If cartonCnt = lastCartonPallet Then
                                            cp_seq += 1
                                            pc_seq = 1
                                            cartonCnt = 0
                                            lastCartonPallet = -1
                                        ElseIf lastCartonPallet = -1 AndAlso cartonCnt Mod CartonPallet = 0 Then
                                            cp_seq += 1
                                            pc_seq = 1
                                            cartonCnt = 0
                                        End If
                                    End If
                                Next
                            Else
                                'If pcsCarton = 0 Then
                                '    uiFun.displayMsg(Me, "", "PCS/ Carton cannot be zero!", Session("gLang"))
                                'ElseIf CartonPallet = 0 Then
                                '    uiFun.displayMsg(Me, "", "Carton/ Pallet cannot be zero!", Session("gLang"))
                                'End If

                                Continue For
                            End If
                        Next

                        sw.WriteLine("DATA_END")

                        sw.Close()
                    End Using
                End Using

                Dim path As String = storingPath & "\" & fileName & ".txt"
                Dim file As System.IO.FileInfo = New System.IO.FileInfo(path)

                If file.Exists Then
                    Response.Clear()
                    Response.AddHeader("Content-Disposition", "attachment; filename=" & file.Name)
                    Response.AddHeader("Content-Length", file.Length.ToString())
                    Response.ContentType = "application/octet-stream"
                    Response.WriteFile(file.FullName)
                End If

            Catch Ex As Exception
                Response.Write(Ex.Message)
            End Try
        Else
            uiFun.displayMsg(Me, "", "No item to export!", Session("gLang"))
        End If
    End Sub

    Protected Sub ddl_BatchChanged(sender As Object, e As EventArgs)
        Dim ddl As DropDownList = sender
        Dim grow As GridViewRow = ddl.NamingContainer

        If CType(grow.FindControl("HOLD_QTY"), TextBox).Text.Trim <> 0 Then
            uiFun.displayMsg(Me, "", "Hold Stock qty. exists! Batch No. cannot be changed!", Session("gLang"))
            ddl.SelectedValue = CType(grow.FindControl("ori_batch"), HiddenField).Value
        End If

        'If ddl.SelectedValue <> "" Then
        '    CType(grow.FindControl("btnHold"), Button).Enabled = True
        'Else
        '    CType(grow.FindControl("btnHold"), Button).Enabled = False
        'End If

        If ar.hasBtnRight("BT_CO_HOLD") Then
            'If ddl.SelectedValue <> "" Then
            '    CType(grow.FindControl("btnHold"), Button).Enabled = True
            '    CType(grow.FindControl("HOLD_QTY"), TextBox).Visible = True
            'End If
        Else
            CType(grow.FindControl("btnHold"), Button).Enabled = False
            CType(grow.FindControl("HOLD_QTY"), TextBox).Visible = False
            CType(grow.FindControl("HOLD_QTY"), TextBox).Text = ""
        End If

        CType(grow.FindControl("ori_batch"), HiddenField).Value = ddl.SelectedValue

    End Sub

    Protected Sub SendCustCOMail(ByRef c_code As String, ByRef st_code As String)
        Dim app_email As String = ""
        Dim mail_title As String = "A New Stock Issue Request has been created by customer."
        Dim mail_body As String = ""

        Dim selectSQL As String = "Select distinct lower(usr_email) as mailadd from wms_user, wms_user_group_alloc where " & _
                                  " wms_user.usr_id = wms_user_group_alloc.usr_id " & _
                                  " and wms_user_group_alloc.grp_code = 'CO_MAIL_GRP' "

        Dim mailDT As New DataTable


        Dim getValSql = "select sto_name as VALUE from wms_storer where storer_code = '" & gU.dbEncode(st_code) & "' "
        Dim lSTO_NAME = DB.getValueFromSQL(getValSql)


        mailDT = gDB.getDataTable(selectSQL)

        If mailDT.Rows.Count > 0 Then
            For i = 0 To mailDT.Rows.Count - 1
                app_email &= mailDT.Rows(i).Item("mailadd").ToString.Trim & ","
            Next
            app_email = Left(app_email, Len(app_email) - 1)
        Else
            app_email = gU.getConfig("AdminEmail")
        End If

        mail_body &= "The following Stock Issue Request has been created:" & vbNewLine
        mail_body &= "  System CO Code: " & c_code & vbNewLine
        'mail_body &= "  Storer Name: " & st_code & vbNewLine
        mail_body &= "  Storer Name: " & lSTO_NAME & vbNewLine
        mail_body &= "  Created Date: " & Now.Date.ToString("dd/MM/yyyy HH:mm:ss")

        Dim mailog As PrgmLog

        If Not System.IO.Directory.Exists(HttpRuntime.Cache("SYSP_LOG_DIR")) Then
            System.IO.Directory.CreateDirectory(HttpRuntime.Cache("SYSP_LOG_DIR"))
        End If

        mailog = New PrgmLog(HttpRuntime.Cache("SYSP_LOG_DIR"), "emailLog" & Now.Date.ToString("ddMMyyyy") & ".txt")

        Call gU.sendEmail(app_email, mail_title, mail_body, mailog)


    End Sub

    Private Function returnBatchNO(ByVal batchPattern As String, ByVal batch_no As String, ByVal manu_date As String, exp_date As String) As String
        Dim returnBatch As String = ""

        Select Case batchPattern
            Case "MANU_DATE"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If manu_date <> "" Then
                        returnBatch = "@B#_M_" & manu_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case "EXP_DATE"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If exp_date <> "" Then
                        returnBatch = "@B#_E_" & exp_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case "EXP_MANU"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If exp_date <> "" Then
                        returnBatch = "@B#_E_" & exp_date.Replace("/", "")
                    ElseIf manu_date <> "" Then
                        returnBatch = "@B#_M_" & manu_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case Else
                returnBatch = batch_no
        End Select

        Return returnBatch

    End Function

    Public Class VendUpdtInfo
        Public PCS_UOM, VOL, WGT As String
        Public PCS_UOM_ID, VOL_ID, WGT_UOM_ID As String
    End Class

    <System.Web.Services.WebMethod()> _
    Public Shared Function updateVendCode(ByVal storer As String, ByVal itemCode As String, ByVal packKey As String, ByVal vendID As String, ByVal vendCode As String) As VendUpdtInfo
        Dim tmpVendInfo As New VendUpdtInfo
        Dim gDB As New GlobalDBFunc
        Dim DB As New DBfunc
        Dim gU As New GeneralUtils
        Dim selectSql As String
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim tmpDt As DataTable


        tmpVendInfo.PCS_UOM_ID = Replace(vendID, "cod_vnd_code", "cod_pcs_uom")
        tmpVendInfo.VOL_ID = Replace(vendID, "cod_vnd_code", "cod_tot_cbm")
        tmpVendInfo.WGT_UOM_ID = Replace(vendID, "cod_vnd_code", "cod_tot_wgt")


        cmdPa = New GlobalDBFunc.DBCmdPara

        selectSql = "select AITM_PCS_PER_PACK, AITM_VOL, AITM_CBM " & _
                    "from wms_alt_vend_item " & _
                    "where imp_code = " & cmdPa.AP(System.Web.HttpContext.Current.Session("IMP_CODE")) & " " & _
                    "and storer_code = " & cmdPa.AP(storer) & " " & _
                    "and itm_code = " & cmdPa.AP(itemCode) & " " & _
                    "and pack_key = " & cmdPa.AP(packKey) & " " & _
                    "and vnd_code = " & cmdPa.AP(vendCode) & " "

        tmpDt = gDB.getDataTable(selectSql, , , , cmdPa)

        If tmpDt.Rows.Count > 0 Then
            'tmpVendInfo.PCS_UOM = tmpDt.Rows(0).Item("AITM_PCS_PER_PACK").ToString.Trim
            tmpVendInfo.WGT = tmpDt.Rows(0).Item("AITM_VOL").ToString.Trim
            tmpVendInfo.VOL = tmpDt.Rows(0).Item("AITM_CBM").ToString.Trim
        Else
            'tmpVendInfo.PCS_UOM = ""
            tmpVendInfo.WGT = ""
            tmpVendInfo.VOL = ""
        End If

        Return tmpVendInfo
    End Function
End Class
