Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient


Partial Class DOMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    'Private moduleAction As String = ""
    Private dl As New DocLink
    Private st As New StockTrans

    Private dt As New DataTable
    Private DDFORMAT As String = gU.getConfig("DDFORMATNo")
    Private exceptionEditList As List(Of String)
    Private dtUOM, dtUOM2 As DataTable

    Private gvCol() As String = {"DOD_DISP_SEQ", "DOD_TICKET_NO", "DOD_PALLET_NO", "DOD_CARTON_NO", "DOD_PACK_NO",
                                 "DOD_BATCH_NO", "DOD_VND_CODE", "DOD_ITM_CODE", "ITM_SKU_NO", "DOD_ITM_DESC", "DOD_CUT_YN", "ITM_DESC",
                                 "DOD_PACK_TYPE", "DOD_PACK_KEY", "DOD_EXPIRY_DATE", "DOD_MANU_DATE", "DOD_QTY", "DOD_UOM",
                                 "DOD_QTY2", "DOD_UOM2", "DOD_PCS_UOM", "DOD_TOTPCS",
                                 "DOD_TOT_WGT", "DOD_TOT_CBM", "DOD_REM"}

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'DDFORMAT = gU.getConfig("DDFORMATNo")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            Exit Sub
        End If

        'moduleAction = Request("moduleAction")

        exceptionEditList = New List(Of String)

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("newWin") = ""
            ViewState("newWin") = Request("newWin")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            If Session("gLang") = "C" Then
                uiFun.load_dropdown(DO_SHIP_MODE, "select colc_code, colc_eng_value, colc_chi_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_chi_value", , Session("gSelectLabel"))
            Else
                uiFun.load_dropdown(DO_SHIP_MODE, "select colc_code, colc_eng_value, colc_chi_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))
            End If
            uiFun.load_dropdown(DO_SHIP_MODE, "select colc_code, colc_eng_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))
            'uiFun.load_dropdown(RT_WH, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))

            Session("DO_CODE") = Server.UrlDecode(Request("DO_CODE"))
            Session("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))
            'ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            btnRefresh.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            If DO_DATE.Text = "" Then
                DO_DATE.Text = Now.Date.ToString(gU.getConfig("DDFORMAT2"))
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Warehouse Issue Ticket Maintenance"
            lbl_DO_CODE.Text = "WWIT Code:"
            lbl_DO_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_DO_CUS_REF_NO.Text = "SIR No.:"
            lbl_DO_DATE.Text = "Date:"
            lbl_DO_ISSUED_BY.Text = "Issued By:"
            lbl_DO_CO_CODE.Text = "WSIR No.:"
            lbl_DO_TARGET_DELDATE.Text = "Target Delivery Date:"
            lbl_DO_CONF_DELDATE.Text = "Confirmed Deli. Date:"
            lbl_DO_CONF_DELTIME.Text = "Confirmed Deli. Time:"
            lbl_DO_PROJECT_NO.Text = "Project:"
            lbl_CUS_CODE.Text = "Customer:"
            lbl_DO_ADDR1.Text = "Address:"
            'lbl_DO_AREA_DEL.Text = "Area"
            'lbl_DO_REGION_DEL.Text = "Region"
            'lbl_DO_COUNTRY_DEL.Text = "Country"
            lbl_DO_CUS_CONT.Text = "Attention:"
            lbl_DO_CUS_CONT_TEL.Text = "Tel No.:"
            lbl_DO_DRIVER.Text = "Driver:"
            lbl_DO_DRIVER_TEL.Text = "Driver Tel.:"
            lbl_DO_VEHICLE_NO.Text = "Vehicle No.:"
            lbl_DO_TOTL_PALLETS.Text = "Total Pallet:"
            lbl_DO_TOTL_CARTONS.Text = "Total Carton:"
            lbl_DO_TOTL_BINS.Text = "Total Bin:"
            lbl_DO_REM.Text = "Remarks:"
            lbl_DO_TRACK_NO.Text = "Track No.:"
            lbl_DO_FTRACK_NO.Text = "Forwarder Tracking No.:"
            lbl_DO_TRANS_TYPE.Text = "Nature of Transaction"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            'New Field
            lbl_DO_CONSIGNEE.Text = "Consignee Name:"
            lbl_DO_CONSIGNEE_ADDR1.Text = "Consignee Address:"
            lbl_DO_SHIP_TO.Text = "Ship to:"
            lbl_DO_SHIP_ADDR1.Text = "Ship Address:"
            lbl_DO_SHIP_MODE.Text = "Ship Mode:"
            lbl_DO_PAY_TERMS.Text = "Pay Terms:"
            lbl_DO_TRADE_TERMS.Text = "Trade Terms:"
            lbl_DO_INV_NO.Text = "Invoice No.:"
            'End here

            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting) "");"
            btnUnPost.OnClientClick = "return confirm(""Are you sure to un-post this record?\r\n(Please save your work before Un-Posting)"");"
            If Session("pagemode") = "N" Then
                DO_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Value = "提貨單維護"
            lbl_DO_CODE.Text = "提貨號碼:"
            lbl_DO_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "儲存庫:"
            lbl_DO_CUS_REF_NO.Text = "文件編號:"
            lbl_DO_DATE.Text = "日期:"
            lbl_DO_ISSUED_BY.Text = "核發者:"
            lbl_DO_CO_CODE.Text = "客戶訂單號碼:"
            lbl_DO_TARGET_DELDATE.Text = "預定運抵日期:"
            lbl_DO_CONF_DELDATE.Text = "確認運抵日期:"
            lbl_DO_CONF_DELTIME.Text = "確認運抵時間:"
            lbl_DO_PROJECT_NO.Text = "項目:"
            lbl_CUS_CODE.Text = "客戶:"
            lbl_DO_ADDR1.Text = "地址:"
            'lbl_DO_AREA_DEL.Text = "地區" 
            'lbl_DO_REGION_DEL.Text = "區域" 
            'lbl_DO_COUNTRY_DEL.Text = "國家" 
            lbl_DO_CUS_CONT.Text = "聯絡人:"
            lbl_DO_CUS_CONT_TEL.Text = "電話號碼:"
            lbl_DO_DRIVER.Text = "司機:"
            lbl_DO_DRIVER_TEL.Text = "司機電話號碼:"
            lbl_DO_VEHICLE_NO.Text = "車輛號碼:"
            lbl_DO_TOTL_PALLETS.Text = "貨板總量:"
            lbl_DO_TOTL_CARTONS.Text = "外箱總量:"
            lbl_DO_TOTL_BINS.Text = "盒總量:"
            lbl_DO_TRACK_NO.Text = "追查編號:"
            lbl_DO_FTRACK_NO.Text = "運送追查編號:"
            lbl_DO_TRANS_TYPE.Text = "貨單性質"
            lbl_DO_REM.Text = "備註:"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"

            REM New Field
            lbl_DO_CONSIGNEE.Text = "Consignee Name:"
            lbl_DO_CONSIGNEE_ADDR1.Text = "Consignee Address:"
            lbl_DO_SHIP_TO.Text = "Ship to:"
            lbl_DO_SHIP_ADDR1.Text = "Ship Address:"
            lbl_DO_SHIP_MODE.Text = "Ship Mode:"
            lbl_DO_PAY_TERMS.Text = "Pay Terms:"
            lbl_DO_TRADE_TERMS.Text = "Trade Terms:"
            lbl_DO_INV_NO.Text = "Invoice No.:"
            ' New Field

            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            selectItemBtn.Text = "選擇物料"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定保存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定保存資料?"");"
            btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnUnPost.OnClientClick = "return confirm(""確定取消發布資料?"");"
            If Session("pagemode") = "N" Then
                DO_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        'RT_TYPE.CssClass = "REQUIRED"
        DO_CO_CODE.CssClass = "REQUIRED"
        REM **********************

        selectItemBtn.Visible = False

        If Session("pagemode") = "N" Then
            'DO_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'DO_CODE.Enabled = False
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            Session("dt") = Nothing

            Session("n_cur_seq") = ""
            Session("DO_CODE") = ""

            Call BindGV()

        Else
            dt = Session("dt")

            If editMode.Value = "V" Then
                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
        End If

        If moduleAction.Value = "SELECTCO" Then
            addCOtoDO()
        End If

        If moduleAction.Value = "RELOADPACK" Then
            reloadPackList()
            recalPackTotal()
        End If

        If moduleAction.Value = "SELECTIM" Then
            addItemtoDO()
        End If

        Dim colIdx_StartWith As Integer = 0

        ar.addColDef("DOD_DISP_SEQ", "dod_disp_seq", colIdx_StartWith)
        ar.addColDef("DOD_TICKET_NO", "DOD_TICKET_NO", colIdx_StartWith)
        ar.addColDef("DOD_PALLET_NO", "dod_pallet_no", colIdx_StartWith)
        ar.addColDef("DOD_CARTON_NO", "dod_carton_no", colIdx_StartWith)
        ar.addColDef("DOD_PACK_NO", "dod_pack_no", colIdx_StartWith)
        ar.addColDef("DOD_VND_CODE", "dod_itm_code", colIdx_StartWith)
        ar.addColDef("DOD_ITM_CODE", "dod_itm_code", colIdx_StartWith)
        ar.addColDef("ITM_SKU_NO", "itm_sku_no", colIdx_StartWith)
        ar.addColDef("DOD_ITM_DESC", "dod_itm_desc", colIdx_StartWith)
        ar.addColDef("DOD_CUT_YN", "dod_cut_yn", colIdx_StartWith)
        ar.addColDef("ITM_DESC", "itm_desc", colIdx_StartWith)
        ar.addColDef("DOD_PACK_TYPE", "dod_pack_type", colIdx_StartWith)
        ar.addColDef("DOD_PACK_KEY", "dod_pack_key", colIdx_StartWith)
        ar.addColDef("DOD_QTY", "dod_qty", colIdx_StartWith)
        ar.addColDef("DOD_UOM", "dod_uom", colIdx_StartWith)
        ar.addColDef("DOD_QTY2", "dod_qty2", colIdx_StartWith)
        ar.addColDef("DOD_UOM2", "dod_uom2", colIdx_StartWith)
        ar.addColDef("DOD_PCS_UOM", "dod_pcs_uom", colIdx_StartWith)
        ar.addColDef("DOD_TOTPCS", "dod_totpcs", colIdx_StartWith)
        ar.addColDef("DOD_TOT_WGT", "dod_tot_wgt", colIdx_StartWith)
        ar.addColDef("DOD_TOT_CBM", "dod_tot_cbm", colIdx_StartWith)
        ar.addColDef("DOD_REM", "dod_rem", colIdx_StartWith)
        'ar.addColDef("DOD_NO_OF_CARTON", "dod_no_of_carton", colIdx_StartWith)

        REM Generate Common Menu
        cm = New CommonMenu("DO", lheader.Value, DO_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "N"
        cm.haveAttachments = "N"
        cm.haveNotes = "N"
        cm.haveTasks = "N"
        cm.haveEmail = "N"
        cm.haveHistory = "N"
        cm.genCM(cmBar)

        REM generate Link Bar
        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "DO", DO_CODE.Text, "../../")

        REM Select RO button
        If DO_CO_CODE.Text <> "" And Session("pagemode") <> "N" Then
            selectCOBtn.Visible = False
        Else
            selectCOBtn.Attributes.Add("onclick", "COLookUp('" & DO_CO_CODE.ClientID & "', '" & DO_CO_CODE.ClientID & "', document.myform." & STORER_CODE.ClientID & ".value);")
        End If

        If DO_CO_CODE.Text <> "" Then
            selectCOBtn.Visible = False
        End If

        REM Select RO button
        If DO_STATUS.Text = "POSTED" Or DO_STATUS.Text = "CANCELLED" Then
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Enabled = True
        Else
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")
            cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);")
            cSBBtn.Enabled = True
        End If
        REM ****************************************************************

        If DO_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue
        ElseIf DO_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
            btnPost.Visible = False
            btnUnPost.Visible = True

            btnPick.Visible = False
            btnUnPick.Visible = False

            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

            exceptionEditList.Add("saveConDate")
            exceptionEditList.Add("DO_CONF_DELTIME")
            exceptionEditList.Add("DO_CONF_DELDATE")
            exceptionEditList.Add("ImageButton3")
            exceptionEditList.Add("CalendarExtender3")
        ElseIf DO_STATUS.Text = "PICKED" Then
            btnPost.Visible = True
            btnPick.Visible = False
            btnUnPick.Visible = True
            btnUnPost.Visible = False
        Else
            If Session("pagemode") = "N" Then btnPick.Visible = False Else btnPick.Visible = True
            btnUnPick.Visible = False
            btnPost.Visible = False
            btnUnPost.Visible = False
        End If

        If ViewState("newWin") = "Y" Then
            ar.sec_write = "N"
            btnBack2.Visible = False
            btnBack.Visible = False

            BtnClose1.Visible = True
            btnClose2.Visible = True
        Else
            btnBack2.Visible = True
            btnBack.Visible = True

            BtnClose1.Visible = False
            btnClose2.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
        'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "DO", "WMS_DELV_ORDER_D")


        If DO_STATUS.Text = "POSTED" Then
            Dim checkGP As String = "select usr_id from wms_user_group_alloc where grp_code = 'ADMIN_GP' and usr_id = '" & Session("usr_id") & "'"
            Dim checkGPDt As DataTable
            checkGPDt = gDB.getDataTable(checkGP)
            If checkGPDt.Rows.Count > 0 Then
                'btnUnPost.Enabled = True
                btnUnPost.Enabled = ar.hasBtnRight("BT_DO_UNPOST")
            Else
                btnUnPost.Enabled = False
            End If

            saveBtn1.Visible = False
            saveBtn2.Visible = False
            saveConDate.Visible = True
            saveConDate.Enabled = True

        End If
        cSBBtn.Enabled = True

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OB_DO','" & Session("imp_code") & "||" & Session("STORER_CODE") & "||" & ViewState("DO_CODE") & "','N');")

        If moduleAction.Value = "SAVEOK" Then
            save()
        End If

        'ar.setFieldCustomize(Me, "OB_DO", STORER_CODE.SelectedValue, "WMS_DELV_ORDER")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "OB_DO", STORER_CODE.SelectedValue, "WMS_DELV_ORDER_D", gvCol)
        'End If
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(saveBtn1)
        sm.RegisterAsyncPostBackControl(saveBtn2)

        sm.RegisterAsyncPostBackControl(selectCOBtn)
        sm.RegisterAsyncPostBackControl(saveConDate)
        sm.RegisterAsyncPostBackControl(CancelBtn)
        sm.RegisterAsyncPostBackControl(cSBBtn)
        sm.RegisterAsyncPostBackControl(btnPick)
        sm.RegisterAsyncPostBackControl(btnUnPick)
        sm.RegisterAsyncPostBackControl(btnPost)
        sm.RegisterAsyncPostBackControl(btnUnPost)
        sm.RegisterAsyncPostBackControl(btnRelease)

        For i = 0 To GridView1.Rows.Count - 1
            'sm.RegisterAsyncPostBackControl(CType(GridView1.Rows(i).FindControl("btnItemSerial"), ImageButton))
        Next

        If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
            'Set labels, attributes, etc
            setGeneralControls()

            'Set field access (hide, readonly, view mode, etc)
            'setPageCtrlAccess()
        End If

        moduleAction.Value = ""

    End Sub

    Private Sub setGeneralControls()
        If DO_CODE_HF.Value <> "" AndAlso DO_STATUS.Text = "NEW" Then
            btnRelease.Visible = True
            btnRelease.OnClientClick = "return confirm('Confirm to release Picking List to PDA?');"
        End If
    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Or ctl.ID = "btnUnPost" Then
            If ctl.ID = "btnUnpost" Then
                CType(ctl, Button).Enabled = ar.hasBtnRight("BT_DO_UNPOST")
            Else
                CType(ctl, Button).Enabled = True
            End If
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

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                'Dim oGridView As GridView = DirectCast(sender, GridView)
                'Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                'REM **********************
                'REM Use for re-create the label to change the Langauge
                'REM Modify Here

                'Call cU.changeGVLabel(oGridViewRow, e, "Seq No.", "編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Tracking No.", "追查編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "外箱編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pack No.", "封裝編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
                'Call cU.changeGVLabel(oGridViewRow, e, "Vendor Code", "供應商代碼")
                'Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
                'Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
                'Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名稱")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pack Type", "封裝形式")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Qty", "數量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "UOM", "單位")
                'Call cU.changeGVLabel(oGridViewRow, e, "Number/ UOM", "單位件數")
                'Call cU.changeGVLabel(oGridViewRow, e, "Total Number", "總件數")
                'Call cU.changeGVLabel(oGridViewRow, e, "Total Weight(kg)", "總重量")
                'Call cU.changeGVLabel(oGridViewRow, e, "Total Volume(cbm)", "總體積(cbm)")
                'Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
                'Call cU.changeGVLabel(oGridViewRow, e, "", "")
                'REM **********************

                'oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub recalPackTotal()

    End Sub

    Protected Sub reloadPackList()
        Dim i, j As Integer
        Dim pi_dt As DataTable
        Dim nDropDown As DropDownList
        Dim addon_item As ListItem
        Dim selItemList As New List(Of String)
        Dim selectedValue As String

        pi_dt = Session("_M_OB_DO_TMP_pi_dt")
        For i = 0 To pi_dt.Rows.Count - 1
            selItemList.Add(pi_dt.Rows(i).Item("PAD_PACK_NO").ToString.Trim)
        Next
        selItemList.Sort()

        If GridView1.Rows.Count > 0 Then
            For j = 0 To GridView1.Rows.Count - 1
                nDropDown = CType(GridView1.Rows(j).FindControl("dod_pack_no"), DropDownList)

                selectedValue = nDropDown.SelectedValue

                nDropDown.Items.Clear()

                addon_item = New ListItem
                addon_item.Text = "SELECT"
                addon_item.Value = ""
                nDropDown.Items.Insert(0, addon_item)

                For i = 0 To selItemList.Count - 1
                    addon_item = New ListItem

                    addon_item.Text = selItemList(i).ToString
                    addon_item.Value = selItemList(i).ToString
                    nDropDown.Items.Insert(nDropDown.Items.Count, addon_item)
                Next

                nDropDown.SelectedValue = selectedValue
            Next
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        Dim pi_dt As DataTable

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                'If Session("gLang") = "E" Then
                '    CType(e.Row.FindControl("gvl_seq"), Label).Text = "Seq No."
                '    CType(e.Row.FindControl("gvl_pallet_no"), Label).Text = "Pallet No."
                '    CType(e.Row.FindControl("gvl_carton_no"), Label).Text = "Carton No."
                '    CType(e.Row.FindControl("gvl_pack_no"), Label).Text = "Pack No."
                '    CType(e.Row.FindControl("gvl_itm_code"), Label).Text = "Item Code"
                '    CType(e.Row.FindControl("gvl_pack_key"), Label).Text = "Pack Key"
                '    CType(e.Row.FindControl("gvl_itm_desc"), Label).Text = "Item Name"
                '    CType(e.Row.FindControl("gvl_pack_type"), Label).Text = "Pack Type"
                '    CType(e.Row.FindControl("gvl_qty"), Label).Text = "Qty"
                '    CType(e.Row.FindControl("gvl_uom"), Label).Text = "UOM"
                '    CType(e.Row.FindControl("gvl_pcs_uom"), Label).Text = "PCS per UOM"
                '    CType(e.Row.FindControl("gvl_totpcs"), Label).Text = "Total PCS"
                '    CType(e.Row.FindControl("gvl_tot_wgt"), Label).Text = "Total Weight"
                '    CType(e.Row.FindControl("gvl_tot_cbm"), Label).Text = "Total CBM"
                '    CType(e.Row.FindControl("gvl_rem"), Label).Text = "Remarks"
                'Else
                '    CType(e.Row.FindControl("gvl_seq"), Label).Text = "编号"
                '    CType(e.Row.FindControl("gvl_pallet_no"), Label).Text = "貨板编号"
                '    CType(e.Row.FindControl("gvl_carton_no"), Label).Text = "外箱编号"
                '    CType(e.Row.FindControl("gvl_pack_no"), Label).Text = "封装编号"
                '    CType(e.Row.FindControl("gvl_itm_code"), Label).Text = "物件号码"
                '    CType(e.Row.FindControl("gvl_pack_key"), Label).Text = "封装内码"
                '    CType(e.Row.FindControl("gvl_itm_desc"), Label).Text = "物件名称"
                '    CType(e.Row.FindControl("gvl_pack_type"), Label).Text = "封装形式"
                '    CType(e.Row.FindControl("gvl_qty"), Label).Text = "数量"
                '    CType(e.Row.FindControl("gvl_uom"), Label).Text = "单位"
                '    CType(e.Row.FindControl("gvl_pcs_uom"), Label).Text = "单位件数"
                '    CType(e.Row.FindControl("gvl_totpcs"), Label).Text = "总件数"
                '    CType(e.Row.FindControl("gvl_tot_wgt"), Label).Text = "总重量"
                '    CType(e.Row.FindControl("gvl_tot_cbm"), Label).Text = "总体积(cbm)"
                '    CType(e.Row.FindControl("gvl_rem"), Label).Text = "备注"
                'End If

                CType(e.Row.FindControl("dod_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_DISP_SEQ").ToString.Trim
                CType(e.Row.FindControl("dod_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("dod_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_CARTON_NO").ToString.Trim

                Dim nDropDown As DropDownList = CType(e.Row.FindControl("dod_pack_no"), DropDownList)
                Dim addon_item As New ListItem
                Dim selItemList As New List(Of String)

                nDropDown.DataSource = Nothing
                nDropDown.DataBind()

                addon_item.Text = "SELECT"
                addon_item.Value = ""
                nDropDown.Items.Insert(0, addon_item)

                pi_dt = Session("_M_OB_DO_TMP_pi_dt")
                For i = 0 To pi_dt.Rows.Count - 1
                    selItemList.Add(pi_dt.Rows(i).Item("PAD_PACK_NO").ToString.Trim)
                Next
                selItemList.Sort()

                For i = 0 To selItemList.Count - 1
                    addon_item = New ListItem

                    addon_item.Text = selItemList(i).ToString
                    addon_item.Value = selItemList(i).ToString
                    nDropDown.Items.Insert(nDropDown.Items.Count, addon_item)
                Next

                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "DOD_PACK_NO").ToString.Trim

                CType(e.Row.FindControl("dod_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_VND_CODE").ToString.Trim
                CType(e.Row.FindControl("dod_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim
                CType(e.Row.FindControl("dod_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("dod_itm_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_ITM_DESC").ToString.Trim
                'CType(e.Row.FindControl("DOD_TROLLEY_ID"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TROLLEY_ID").ToString.Trim

                If Session("gLang") = "E" Then
                    uiFun.load_dropdown(CType(e.Row.FindControl("dod_pack_type"), DropDownList), "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'PACK_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(CType(e.Row.FindControl("dod_pack_type"), DropDownList), "select colc_code, colc_chi_value from wms_col_code where colc_tabcol = 'PACK_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_CHI_VALUE", , Session("gSelectLabel"))
                End If

                CType(e.Row.FindControl("dod_pack_type"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "DOD_PACK_TYPE").ToString.Trim

                'uiFun.load_dropdown(CType(e.Row.FindControl("DOD_BATCH_NO"), DropDownList), "select dc_date_code from wms_date_code order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("DOD_BATCH_NO"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "DOD_BATCH_NO").ToString.Trim

                CType(e.Row.FindControl("DOD_BATCH_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_BATCH_NO").ToString.Trim

                CType(e.Row.FindControl("DOD_EXPIRY_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_MANU_DATE").ToString.Trim

                CType(e.Row.FindControl("dod_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_QTY").ToString.Trim

                'Dim n1DropDown As DropDownList = CType(e.Row.FindControl("dod_uom"), DropDownList)
                'uiFun.load_dropdown(n1DropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'n1DropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim
                uiFun.load_dropdown(CType(e.Row.FindControl("dod_uom"), DropDownList), dtUOM, "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim)

                CType(e.Row.FindControl("dod_uom"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim

                CType(e.Row.FindControl("dod_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_QTY2").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "DOD_CUT_YN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("dod_cut_yn"), CheckBox).Checked = True
                Else
                    CType(e.Row.FindControl("dod_cut_yn"), CheckBox).Checked = False
                End If

                'n1DropDown = CType(e.Row.FindControl("dod_uom2"), DropDownList)
                'uiFun.load_dropdown(n1DropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'n1DropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom2").ToString.Trim
                uiFun.load_dropdown(CType(e.Row.FindControl("dod_uom2"), DropDownList), dtUOM2, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "dod_uom2").ToString.Trim)


                CType(e.Row.FindControl("dod_uom2"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom2").ToString.Trim

                CType(e.Row.FindControl("dod_pcs_uom"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PCS_UOM").ToString.Trim
                CType(e.Row.FindControl("dod_totpcs"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TOTPCS").ToString.Trim
                CType(e.Row.FindControl("dod_tot_wgt"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TOT_WGT").ToString.Trim
                CType(e.Row.FindControl("dod_tot_cbm"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TOT_CBM").ToString.Trim
                CType(e.Row.FindControl("dod_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_REM").ToString.Trim
                CType(e.Row.FindControl("DOD_TICKET_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TICKET_NO").ToString.Trim

                CType(e.Row.FindControl("dod_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("dod_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("dod_pcs_uom"), TextBox).ClientID & ".value * this.value")
                CType(e.Row.FindControl("dod_pcs_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("dod_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("dod_qty"), TextBox).ClientID & ".value * this.value")

                'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" AndAlso DataBinder.Eval(e.Row.DataItem, "DOD_NO_OF_CARTON").ToString.Trim = "" Then
                '    CType(e.Row.FindControl("dod_no_of_carton"), TextBox).Text = 1
                'Else
                '    CType(e.Row.FindControl("dod_no_of_carton"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "DOD_NO_OF_CARTON").ToString.Trim)
                'End If
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

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                End If

                ar.hideGVForStorer(GridView1, STORER_CODE.Text, "DO", "WMS_DELV_ORDER_D", e)

        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "
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

            If Session("n_cur_seq") = "" Then
                Session("n_cur_seq") = next_seq_no
            Else
                temp_seq_no = CInt(Session("n_cur_seq")) + 1
                Session("n_cur_seq") = temp_seq_no.ToString
            End If


            dt.Rows.Add()
            rows_count = dt.Rows.Count

            REM **********************
            REM Modify Here
            dt.Rows(rows_count - 1).Item("dod_seq") = Session("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            Session("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        If dt.Rows(e.RowIndex).Item("mFlag") = "N" Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
                dt.Rows(e.RowIndex).Delete()

                dt.AcceptChanges()

                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
        Else
            dt.Rows(e.RowIndex).Item("mFlag") = "D"
            GridView1.Rows(e.RowIndex).Visible = False
            dt.AcceptChanges()
        End If
        'Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))

    End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        'If status = picked, should check qty like post
        If DO_STATUS.Text.Trim = "PICKED" OrElse flag = "P" Then
            flag = "Y"
        End If

        If DO_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If DO_CO_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_CO_CODE.Text & " cannot be empty! Please select CO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_CO_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'selectSql = "select count(*) as value " & _
        '            "from WMS_CUST_ORDER " & _
        '            "where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
        '            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
        '            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text.Trim) & "' "

        'If DB.getValueFromSQL(selectSql) = "0" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsgNew(updtPnlAlert, "", "Customer Order (" & DO_CO_CODE.Text & ") is not found!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If DO_DATE.Text.Trim <> "" And Not gU.isValidDate(DO_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid date, " & lbl_DO_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的日期, " & lbl_DO_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If DO_TARGET_DELDATE.Text.Trim <> "" And Not gU.isValidDate(DO_TARGET_DELDATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid date, " & lbl_DO_TARGET_DELDATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的日期, " & lbl_DO_TARGET_DELDATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If DO_CONF_DELDATE.Text.Trim <> "" And Not gU.isValidDate(DO_CONF_DELDATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid date, " & lbl_DO_CONF_DELDATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的日期, " & lbl_DO_CONF_DELDATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If DO_CONF_DELTIME.Text.Trim <> "" And Not gU.isTime(DO_CONF_DELTIME.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid time, " & lbl_DO_CONF_DELTIME.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的時間, " & lbl_DO_CONF_DELTIME.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(DO_TOTL_PALLETS.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, " & lbl_DO_TOTL_PALLETS.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, " & lbl_DO_TOTL_PALLETS.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(DO_TOTL_CARTONS.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, " & lbl_DO_TOTL_CARTONS.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, " & lbl_DO_TOTL_CARTONS.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(DO_TOTL_BINS.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, " & lbl_DO_TOTL_BINS.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, " & lbl_DO_TOTL_BINS.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If DO_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Too many characters, maximum length of " & lbl_DO_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "字數太多, " & lbl_DO_REM.Text & "最多只限200字!", Session("gLang"))
            End If
            Return False
        End If

        'If DO_TRACK_NO.Text.Trim <> "" Then
        '    If hide_DO_TRACK_NO.Value.ToString.Trim <> DO_TRACK_NO.Text.Trim Then
        '        selectSql = "Select DO_CODE from WMS_DELV_ORDER where DO_STATUS<>'CANCELLED' AND IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND DO_TRACK_NO='" & gU.dbEncode(DO_TRACK_NO.Text.Trim) & "'"

        '        Dim seldt As New DataTable

        '        seldt = gDB.getDataTable(selectSql)

        '        If seldt.Rows.Count > 0 Then
        '            uiFun.displayMsgNew(updtPnlAlert, "", "Tracking no. " & DO_TRACK_NO.Text.Trim & " already in use in DO " & seldt.Rows(0).Item("DO_CODE").ToString.Trim, Session("gLang"))
        '            Return False
        '        End If
        '    End If
        'End If



        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then
                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid whole number, Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的整數, 數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_pcs_uom"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Number/ UOM!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 單位件數!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_qty2"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid whole number, Qty 2!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的整數, 數量2!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_totpcs"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Total Number!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 總件數!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_tot_wgt"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Total Weight!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 總重量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("dod_tot_cbm"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Total CBM!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 總體積(cbm)!", Session("gLang"))
                        End If
                        Return False
                    End If
                End If
            Next
        End If

        If Not isValidPLQty(flag) Then
            Return False
        End If

        'If Not isValidItems() Then
        'Return False
        'End If

        Return True

    End Function

    Private Function isValidItems() As Boolean
        Dim i As Integer

        For i = 0 To GridView1.Rows.Count - 1
            If Not DB.chkItmExist(IMP_CODE.Value.Trim, STORER_CODE.SelectedValue, CType(GridView1.Rows(i).FindControl("dod_itm_code"), Label).Text, CType(GridView1.Rows(i).FindControl("dod_pack_key"), Label).Text) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Item code is not existed in item master : " & CType(GridView1.Rows(i).FindControl("dod_itm_code"), Label).Text & "-" & CType(GridView1.Rows(i).FindControl("dod_pack_key"), Label).Text & "!", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", "物件號碼不存在系統中 : " & CType(GridView1.Rows(i).FindControl("dod_itm_code"), Label).Text & "-" & CType(GridView1.Rows(i).FindControl("dod_pack_key"), Label).Text & "!", Session("gLang"))
                End If
                Return False
            End If
        Next

        Return True
    End Function

    Private Function isValidPLQty(Optional ByVal postFlag As String = "") As Boolean
        Dim pl_dt As DataTable
        Dim dtlQtyDict, plQtyDict As Dictionary(Of String, Double)
        Dim i As Integer
        Dim itmKey As String
        Dim keys As Dictionary(Of String, Double).KeyCollection
        Dim itmArray As String()
        Dim javaStr As String

        dtlQtyDict = New Dictionary(Of String, Double)

        For i = 0 To GridView1.Rows.Count - 1
            If GridView1.Rows(i).Visible Then
                itmKey = CType(GridView1.Rows(i).FindControl("dod_itm_code"), Label).Text.Trim & "#_#" & CType(GridView1.Rows(i).FindControl("dod_pack_key"), Label).Text.Trim

                If dtlQtyDict.ContainsKey(itmKey) Then
                    dtlQtyDict.Item(itmKey) = dtlQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0))
                Else
                    dtlQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0)))
                End If
            End If
        Next

        pl_dt = Session("_M_OB_DO_TMP_pl_dt")

        plQtyDict = New Dictionary(Of String, Double)

        For i = 0 To pl_dt.Rows.Count - 1
            If pl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                'Check PL Qty for each row
                If gU.decodeEmptyCdbl(pl_dt.Rows(i).Item("pld_item_qty").ToString, 0) > gU.decodeEmptyCdbl(pl_dt.Rows(i).Item("pld_foi_qty").ToString, 0) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Picked qty is greater than FFI qty, " & pl_dt.Rows(i).Item("pld_item_no").ToString.Trim & "-" & pl_dt.Rows(i).Item("pld_pack_key") & ", please change the qty in Picking List!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量大於預取貨量, " & pl_dt.Rows(i).Item("pld_item_no").ToString.Trim & "-" & pl_dt.Rows(i).Item("pld_pack_key") & ", 請修正取貨單內的數量!", Session("gLang"))
                    End If

                    Return False
                End If

                itmKey = pl_dt.Rows(i).Item("pld_item_no").ToString.Trim & "#_#" & pl_dt.Rows(i).Item("pld_pack_key").ToString.Trim

                If plQtyDict.ContainsKey(itmKey) Then
                    plQtyDict.Item(itmKey) = plQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_item_qty"), 0), 0))
                Else
                    plQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_item_qty"), 0), 0)))
                End If
            End If
        Next

        If postFlag = "Y" Then
            For i = 0 To pl_dt.Rows.Count - 1
                If gU.decodeEmptyCdbl(pl_dt.Rows(i).Item("pld_item_qty"), 0) <= 0 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Picked qty must be greater than zero! " & pl_dt.Rows(i).Item("pld_item_no").ToString.Trim & "-" & pl_dt.Rows(i).Item("pld_pack_key") & "", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量不能為零! " & pl_dt.Rows(i).Item("pld_item_no").ToString.Trim & "-" & pl_dt.Rows(i).Item("pld_pack_key") & "", Session("gLang"))
                    End If
                    Return False
                End If

                If pl_dt.Rows(i).Item("pld_loc").ToString.Trim = "" And gU.decodeEmptyCdbl(pl_dt.Rows(i).Item("pld_item_qty"), 0) >= 0 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Pick list location cannot be empty! Please check pick list.", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "取貨位置不能空白! 請檢查取貨單", Session("gLang"))
                    End If
                    Return False
                    'ElseIf pl_dt.Rows(i).Item("pld_batch_no").ToString.Trim = "" And gU.decodeEmptyCdbl(pl_dt.Rows(i).Item("pld_item_qty"), 0) >= 0 Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsgNew(updtPnlAlert, "", "Pick list batch no. cannot be empty! Please check Pick list.", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsgNew(updtPnlAlert, "", "Date Code不能空白! 請檢查取貨單", Session("gLang"))
                    '    End If
                    '    Return False
                End If
            Next
        End If

        If postFlag = "Y" Then
            keys = dtlQtyDict.Keys

            For i = 0 To keys.Count - 1
                If plQtyDict.ContainsKey(keys(i)) Then
                    If dtlQtyDict.Item(keys(i)) < plQtyDict.Item(keys(i)) Then
                        itmArray = Split(keys(i), "#_#")
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Item qty is less than picked qty, " & itmArray(0) & "-" & itmArray(1) & ", please change the picked qty in Picking List!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "貨單數量少於實取貨量, " & itmArray(0) & "-" & itmArray(1) & ", 請修正取貨單內的數量!", Session("gLang"))
                        End If

                        Return False
                    End If
                Else
                    itmArray = Split(keys(i), "#_#")
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "There is no Picking List for item, " & itmArray(0) & "-" & itmArray(1) & ", please check the Picking List!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "在取貨單中找不到貨品, " & itmArray(0) & "-" & itmArray(1) & ", 請檢楂取貨單!", Session("gLang"))
                    End If

                    Return False
                End If
            Next

        ElseIf moduleAction.Value <> "SAVEOK" And postFlag <> "Y" Then
            keys = dtlQtyDict.Keys

            For i = 0 To keys.Count - 1
                If plQtyDict.ContainsKey(keys(i)) Then
                    If dtlQtyDict.Item(keys(i)) < plQtyDict.Item(keys(i)) Then
                        itmArray = Split(keys(i), "#_#")
                        'If Session("gLang") = "E" Then
                        '    uiFun.displayMsgNew(updtPnlAlert, "", "Over picked qty, " & itmArray(0) & "-" & itmArray(1) & ", please check picking list!", Session("gLang"))
                        'Else
                        '    uiFun.displayMsgNew(updtPnlAlert, "", "取貨過量, " & itmArray(0) & "-" & itmArray(1) & ", 請檢查取貨單!", Session("gLang"))
                        'End If
                        If Session("gLang") = "E" Then
                            javaStr = "if(confirm(""Item qty is less than picked qty, " & itmArray(0) & "-" & itmArray(1) & ", confirm to proceed?"")) saveok();"
                        Else
                            javaStr = "if(confirm(""貨單數量少於取貨量, " & itmArray(0) & "-" & itmArray(1) & ", 確定輸入?"")) saveok();"
                        End If
                        If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                            Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                        End If

                        Return False
                    End If
                Else
                    itmArray = Split(keys(i), "#_#")
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "There is no Picking List for item, " & itmArray(0) & "-" & itmArray(1) & ", please check the Picking List!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "在取貨單中找不到貨品, " & itmArray(0) & "-" & itmArray(1) & ", 請檢楂取貨單!", Session("gLang"))
                    End If

                    Return False
                End If
            Next
        End If

        Return True
    End Function

    Private Function isValidPost() As Boolean
        Dim selectSql As String
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim tmpDt As DataTable
        Dim alertMsg As String
        Dim i As Integer

        Return True

        cmdPa = New GlobalDBFunc.DBCmdPara

        selectSql = "select delv.imp_code, " & _
                        "delv.storer_code, " & _
                        "delv.pld_item_no as itm_code, " & _
                        "i.itm_name, " & _
                        "delv.pld_pack_key as pack_key, " & _
                        "delv.pld_pallet_no as pallet_no, " & _
                        "delv.pld_batch_no as batch_no, " & _
                        "delv.pld_item_qty, " & _
                        "case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end as hold_qty, " & _
                        "ISNULL(case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) as itm_balance, " & _
                        "ISNULL(case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) - " & _
                            "ISNULL(case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end), " & _
                            "0) as avail_qty " & _
                    "from ( " & _
                        "select imp_code, " & _
                            "storer_code, " & _
                            "pld_item_no, " & _
                            "pld_pack_key, " & _
                            "ISNULL(pld_pallet_no, '000') as pld_pallet_no, " & _
                            "pld_batch_no, " & _
                            "sum(pld_item_qty) as pld_item_qty " & _
                        "from wms_do_picklist_d d " & _
                        "where imp_code = " & cmdPa.AP(Session("IMP_CODE")) & " " & _
                        "and storer_code = " & cmdPa.AP(STORER_CODE.SelectedValue) & " " & _
                        "and DO_CODE = " & cmdPa.AP(DO_CODE.Text.Trim) & " " & _
                        "group by imp_code, storer_code, pld_item_no, pld_pack_key, ISNULL(pld_pallet_no, '000'), pld_batch_no) delv, wms_item i, " & _
                        "( " & _
                        "select h.imp_code, " & _
                            "h.storer_code, " & _
                            "d.COD_ITM_CODE, " & _
                            "d.COD_PACK_KEY, " & _
                            "ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, " & _
                            "d.COD_BATCH_NO, " & _
                            "sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                        "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                        "where h.imp_code = d.imp_code " & _
                        "and h.storer_code = d.storer_code " & _
                        "and h.co_code = d.co_code " & _
                        "and h.cod_seq = d.cod_seq " & _
                        "and c.imp_code = d.imp_code " & _
                        "and c.storer_code = d.storer_code " & _
                        "and c.co_code = d.co_code " & _
                        "and h.coh_status <> 'RELEASE' " & _
                        "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                        "and c.co_code <> " & cmdPa.AP(DO_CO_CODE.Text.Trim) & " " & _
                        "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold, " & _
                        "( " & _
                        "select h.imp_code, " & _
                            "h.storer_code, " & _
                            "d.COD_ITM_CODE, " & _
                            "d.COD_PACK_KEY, " & _
                            "ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, " & _
                            "sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                        "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                        "where h.imp_code = d.imp_code " & _
                        "and h.storer_code = d.storer_code " & _
                        "and h.co_code = d.co_code " & _
                        "and h.cod_seq = d.cod_seq " & _
                        "and c.imp_code = d.imp_code " & _
                        "and c.storer_code = d.storer_code " & _
                        "and c.co_code = d.co_code " & _
                        "and h.coh_status <> 'RELEASE' " & _
                        "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                        "and c.co_code <> " & cmdPa.AP(DO_CO_CODE.Text.Trim) & " " & _
                        "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000')) hold2, " & _
                        "( " & _
                        "select l.imp_code, " & _
                            "l.storer_code, " & _
                            "l.ITM_CODE, " & _
                            "l.PACK_KEY, " & _
                            "ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " & _
                            "l.ILOC_BATCH_NO, " & _
                            "sum(l.ILOC_BAL_QTY) as BAL_QTY " & _
                        "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " & _
                        "where ISNULL(ILOC_BAL_QTY, 0) > 0 " & _
                        "and l.IMP_CODE = a.IMP_CODE " & _
                        "and l.ILOC_WH = a.WH_CODE " & _
                        "and l.ILOC_FLOOR = a.FL_NUM " & _
                        "and l.ILOC_AREA = a.AR_CODE " & _
                        "and ISNULL(a.AR_DAMAGE_YN, ' ') <> 'Y' " & _
                        "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000'), l.ILOC_BATCH_NO) bal, " & _
                        "( " & _
                        "select l.imp_code, " & _
                            "l.storer_code, " & _
                            "l.ITM_CODE, " & _
                            "l.PACK_KEY, " & _
                            "ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " & _
                            "sum(l.ILOC_BAL_QTY) as BAL_QTY " & _
                        "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " & _
                        "where ISNULL(l.ILOC_BAL_QTY, 0) > 0 " & _
                        "and l.IMP_CODE = a.IMP_CODE " & _
                        "and l.ILOC_WH = a.WH_CODE " & _
                        "and l.ILOC_FLOOR = a.FL_NUM " & _
                        "and l.ILOC_AREA = a.AR_CODE " & _
                        "and ISNULL(a.AR_DAMAGE_YN, ' ') <> 'Y' " & _
                        "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000')) bal2 " & _
                    "where i.imp_code = delv.imp_code " & _
                    "and i.storer_code = delv.storer_code " & _
                    "and i.itm_code = delv.pld_item_no " & _
                    "and i.pack_key = delv.pld_pack_key " & _
                    "and delv.imp_code = hold.imp_code(+) " & _
                    "and delv.storer_code = hold.storer_code(+) " & _
                    "and delv.pld_item_no = hold.cod_itm_code(+) " & _
                    "and delv.pld_pack_key = hold.cod_pack_key(+) " & _
                    "and delv.pld_pallet_no = hold.cod_pallet_no(+) " & _
                    "and delv.pld_batch_no = hold.cod_batch_no(+) " & _
                    "and delv.imp_code = hold2.imp_code(+) " & _
                    "and delv.storer_code = hold2.storer_code(+) " & _
                    "and delv.pld_item_no = hold2.cod_itm_code(+) " & _
                    "and delv.pld_pack_key = hold2.cod_pack_key(+) " & _
                    "and delv.pld_pallet_no = hold2.cod_pallet_no(+) " & _
                    "and delv.imp_code = bal.imp_code(+) " & _
                    "and delv.storer_code = bal.storer_code(+) " & _
                    "and delv.pld_item_no = bal.ITM_CODE(+) " & _
                    "and delv.pld_pack_key = bal.PACK_KEY(+) " & _
                    "and delv.pld_pallet_no = bal.iloc_pallet_no(+) " & _
                    "and delv.pld_batch_no = bal.iloc_batch_no(+) " & _
                    "and delv.imp_code = bal2.imp_code(+) " & _
                    "and delv.storer_code = bal2.storer_code(+) " & _
                    "and delv.pld_item_no = bal2.ITM_CODE(+) " & _
                    "and delv.pld_pack_key = bal2.PACK_KEY(+) " & _
                    "and delv.pld_pallet_no = bal2.iloc_pallet_no(+) " & _
                    "and ISNULL(Case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) - " & _
                        "ISNULL(case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end, " & _
                        "0) < delv.pld_item_qty " & _
                    "order by delv.pld_item_no "

        tmpDt = gDB.getDataTable(selectSql, , , , cmdPa)

        If tmpDt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                alertMsg = "Picked qty is greater than available qty, please check the stock balance of below items:"
            Else
                alertMsg = "實取貨量大於可取貨量, 請檢查以下貨物的可取存量:"
            End If

            For i = 0 To tmpDt.Rows.Count - 1
                alertMsg = alertMsg & vbNewLine & _
                           "Item: " & tmpDt.Rows(i).Item("itm_code").ToString.Trim & ", Picked: " & tmpDt.Rows(i).Item("pld_item_qty").ToString.Trim & ", Available: " & tmpDt.Rows(i).Item("avail_qty").ToString.Trim
            Next

            'alertMsg = "test"

            uiFun.displayMsgNew(Me, "", alertMsg, Session("gLang"))

            Return False
        End If


        cmdPa = New GlobalDBFunc.DBCmdPara

        selectSql = "select d.pld_item_no, d.pld_loc, ISNULL(d.pld_item_qty, 0) as pick_qty, ISNULL(l.iloc_bal_qty, 0) as bal_qty " & _
                    "from wms_do_picklist_d d, wms_item_loc_bal l " & _
                    "where d.imp_code = l.imp_code " & _
                    "and d.storer_code = l.storer_code " & _
                    "and d.pld_item_no = l.itm_code " & _
                    "and d.pld_pack_key = l.pack_key " & _
                    "and ISNULL(d.pld_pallet_no, '000') = ISNULL(l.iloc_pallet_no, '000') " & _
                    "and ISNULL(d.pld_batch_no, ' ') = ISNULL(l.iloc_batch_no, ' ') " & _
                    "and d.pld_loc = l.iloc_loc " & _
                    "and ISNULL(d.pld_item_qty, 0) > ISNULL(l.iloc_bal_qty, 0) " & _
                    "and d.imp_code = " & cmdPa.AP(Session("IMP_CODE")) & " " & _
                    "and d.storer_code = " & cmdPa.AP(STORER_CODE.SelectedValue) & " " & _
                    "and d.do_code = " & cmdPa.AP(DO_CODE.Text.Trim) & " " & _
                    "order by d.pld_item_no, d.pld_loc "

        tmpDt = gDB.getDataTable(selectSql, , , , cmdPa)

        If tmpDt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                alertMsg = "Picked qty is greater than balance qty, please check the stock balance of below items:"
            Else
                alertMsg = "實取貨量大於貨存量, 請檢查以下貨物的存量:"
            End If

            For i = 0 To tmpDt.Rows.Count - 1
                alertMsg = alertMsg & vbNewLine & _
                           "Item: " & tmpDt.Rows(i).Item("pld_item_no").ToString.Trim & ", Location: " & tmpDt.Rows(i).Item("pld_loc").ToString.Trim & ", Picked: " & tmpDt.Rows(i).Item("pick_qty").ToString.Trim & ", Balance: " & tmpDt.Rows(i).Item("bal_qty").ToString.Trim
            Next

            uiFun.displayMsgNew(Me, "", alertMsg, Session("gLang"))

            Return False

        End If

        Return True
    End Function

    Protected Function save(Optional ByVal flag As String = "") As Boolean
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim updateSql As String
        Dim itemSQL As String = ""
        Dim gConn As SqlConnection
        Dim pi_dt, pl_dt As DataTable
        Dim i, j As Integer
        Dim selectSql As String
        Dim mplDt As DataTable
        Dim plSeq As Integer
        Dim nextNo As String = ""

        If validateAll(flag) Then
            pi_dt = Session("_M_OB_DO_TMP_pi_dt")
            pl_dt = Session("_M_OB_DO_TMP_pl_dt")

            For i = pi_dt.Rows.Count - 1 To 0 Step -1
                If pi_dt.Rows(i).Item("PAD_PACK_NO").ToString.Trim = "" Then
                    pi_dt.Rows(i).Delete()
                End If
            Next
            pi_dt.AcceptChanges()

            delPLZeroRow(pl_dt)

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("DO", gConn, transaction)
                    REM **********************

                    'sql_string = "select count(*) from WMS_DELV_ORDER " & _
                    '            "where IMP_CODE = '" & Session("imp_code") & "' " & _
                    '            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    '            "and DO_CODE = '" & gU.dbEncode(nextNo) & "'"

                    'If CInt(DB.getValueFromSQL(sql_string)) > 0 Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsgNew(updtPnlAlert, "", "Duplicate key! " & lbl_DO_CODE.Text & "(" & DO_CODE.Text.Trim & ") is already existed!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsgNew(updtPnlAlert, "", "资料重复! " & lbl_DO_CODE.Text & "(" & DO_CODE.Text.Trim & ")已经存在!", Session("gLang"))
                    '    End If
                    '    Exit Sub
                    'End If

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "insert into WMS_DELV_ORDER " & _
                    "(IMP_CODE, STORER_CODE, DO_CODE, DO_STATUS, DO_ISSUED_BY, DO_CO_CODE, DO_CUS_REF_NO, DO_PROJECT_NO, DO_DATE, DO_TARGET_DELDATE, " & _
                    "DO_CONF_DELDATE, DO_CONF_DELTIME, CUS_CODE, CUS_NAME, DO_ADDR1, DO_ADDR2, DO_ADDR3, DO_AREA_DEL, DO_REGION_DEL, " & _
                    "DO_COUNTRY_DEL, DO_CUS_CONT, DO_CUS_CONT_TEL, DO_DRIVER, DO_DRIVER_TEL, DO_VEHICLE_NO, DO_TOTL_PALLETS, " & _
                    "DO_TOTL_CARTONS, DO_TOTL_BINS, DO_REM, DO_track_no, DO_FTRACK_NO, " & _
                    "DO_CONSIGNEE, DO_CONSIGNEE_ADDR1, DO_CONSIGNEE_ADDR2, DO_CONSIGNEE_ADDR3, DO_SHIP_TO, DO_SHIP_ADDR1,DO_SHIP_ADDR2,DO_SHIP_ADDR3, " & _
                    "DO_PAY_TERMS, DO_TRADE_TERMS, DO_SHIP_MODE, DO_INV_NO, DO_PACK_LABEL_1, DO_PACK_LABEL_2, DO_PACK_LABEL_3, DO_PACK_LABEL_QTY_1, " & _
                    "DO_PACK_LABEL_QTY_2, DO_PACK_LABEL_QTY_3, DO_TRANS_TYPE,DO_EDI_SIR_NO, DO_TROLLEY_ID, DO_DRUM_ID, " & _
                     "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                    "values ( " & _
                    "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', '" & gU.dbEncode(DO_STATUS.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_ISSUED_BY.Text.Trim) & "', '" & gU.dbEncode(DO_CO_CODE.Text.Trim) & "', '" & gU.dbEncode(DO_CUS_REF_NO.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_PROJECT_NO.SelectedValue) & "', " & gU.convdbDate(gU.dbEncode(DO_DATE.Text.Trim)) & ", " & gU.convdbDate(gU.dbEncode(DO_TARGET_DELDATE.Text.Trim)) & ", " & _
                    gU.convdbDate(gU.dbEncode(DO_CONF_DELDATE.Text.Trim)) & ", "

                    If DO_CONF_DELDATE.Text.Trim <> "" AndAlso DO_CONF_DELTIME.Text.Trim <> "" Then
                        sql_string = sql_string & _
                            "to_date('" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', '" & gU.getConfig("DDFORMAT") & " hh24:mi'), "
                    Else
                        sql_string = sql_string & "null, "
                    End If

                    sql_string = sql_string & _
                    "'" & gU.dbEncode(CUS_CODE.SelectedValue) & "', '" & gU.dbEncode(CUS_NAME.Text.Trim) & "', '" & gU.dbEncode(DO_ADDR1.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_ADDR2.Text.Trim) & "', '" & gU.dbEncode(DO_ADDR3.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_AREA_DEL.Text.Trim) & "', '" & gU.dbEncode(DO_REGION_DEL.Text.Trim) & "', '" & gU.dbEncode(DO_COUNTRY_DEL.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_CUS_CONT.Text.Trim) & "', '" & gU.dbEncode(DO_CUS_CONT_TEL.Text.Trim) & "', '" & gU.dbEncode(DO_DRIVER.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_DRIVER_TEL.Text.Trim) & "', '" & gU.dbEncode(DO_VEHICLE_NO.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_PALLETS.Text.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_CARTONS.Text.Trim, "0")) & "', " & _
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_BINS.Text.Trim, "0")) & "', '" & gU.dbEncode(DO_REM.Text.Trim) & "', '" & gU.dbEncode(DO_TRACK_NO.Text.Trim) & "', '" & gU.dbEncode(DO_FTRACK_NO.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_CONSIGNEE.Text.Trim) & "', " & "'" & gU.dbEncode(DO_CONSIGNEE_ADDR1.Text.Trim) & "'," & "'" & gU.dbEncode(DO_CONSIGNEE_ADDR2.Text.Trim) & "', " & "'" & gU.dbEncode(DO_CONSIGNEE_ADDR3.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_SHIP_TO.Text.Trim) & "', " & "'" & gU.dbEncode(DO_SHIP_ADDR1.Text.Trim) & "', " & "'" & gU.dbEncode(DO_SHIP_ADDR2.Text.Trim) & "', " & "'" & gU.dbEncode(DO_SHIP_ADDR3.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_PAY_TERMS.Text.Trim) & "', " & "'" & gU.dbEncode(DO_TRADE_TERMS.Text.Trim) & "', " & "'" & gU.dbEncode(DO_SHIP_MODE.SelectedValue) & "', " & "'" & gU.dbEncode(DO_INV_NO.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_PACK_LABEL_1.Value.Trim) & "', " & "'" & gU.dbEncode(DO_PACK_LABEL_2.Value.Trim) & "', " & "'" & gU.dbEncode(DO_PACK_LABEL_3.Value.Trim) & "', " & _
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_1.Value.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_2.Value.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_3.Value.Trim, "0")) & "', " & _
                    "'" & gU.dbEncode(DO_TRANS_TYPE.SelectedValue) & "', " & "'" & gU.dbEncode(DO_EDI_SIR_NO.Text.Trim) & "', " & _
                    "'" & gU.dbEncode(DO_TROLLEY_ID.Text.Trim) & "', " & "'" & gU.dbEncode(DO_DRUM_ID.Text.Trim) & "', " & _
                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '"'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & nextNo & "', " & _
                    REM **********************

                    REM Generate Document Link 
                    'dl.genDocLink("DO", DO_CODE.Text.Trim, "CO", DO_CO_CODE.text, STORER_CODE.SelectedValue, gConn, transaction)
                    dl.genDocLink("DO", nextNo, "CO", DO_CO_CODE.Text, STORER_CODE.SelectedValue, gConn, transaction)

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                        uiFun.reOrderDetails(dt, "dod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into WMS_DELV_ORDER_D " & _
                                            "(IMP_CODE, STORER_CODE, DO_CODE, DOD_SEQ, DOD_DISP_SEQ, DOD_PALLET_NO, DOD_CARTON_NO, DOD_PACK_NO, DOD_ITM_CODE, DOD_PACK_KEY, " & _
                                            "DOD_ITM_DESC, DOD_PACK_TYPE, DOD_QTY, DOD_UOM, DOD_CUT_YN, DOD_QTY2, DOD_UOM2, DOD_PCS_UOM, DOD_TOTPCS, DOD_TOT_WGT, DOD_TOT_CBM, DOD_REM, DOD_TICKET_NO, DOD_VND_CODE, " & _
                                            "DOD_BATCH_NO, DOD_EXPIRY_DATE, DOD_MANU_DATE, " & _
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                            "values " & _
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_cut_yn").ToString.Trim, "N")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " & _
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " & _
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " & _
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                             "N'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    '"('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.Text) & "', '" & nextNo & "', " & _
                            End Select
                            REM **********************

                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If

                    For Each rows As DataRow In pi_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"
                                itemSQL = "insert into wms_do_packing_d " & _
                                        "(imp_code,storer_code,do_code," & _
                                        "pad_pack_no,pad_pallet_no,pad_pack_type," & _
                                        "pad_totl_packs, pad_uom, pad_carton_no, " & _
                                        "pad_pack_key, pad_itm_code, pad_vnd_code, " & _
                                        "pad_batch_no, pad_qty, pad_length, " & _
                                        "pad_width, pad_height, pad_ref_no, " & _
                                        "pad_pack_by, pad_display_seq, pad_net_weight, pad_gross_weight, pad_min_packing, pad_origin,pad_pack_size, PAD_QTY_PER_CTN, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values " & _
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, "0")) & "', " & _
                                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                '"('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.Text) & "', '" & nextNo & "', " & _
                        End Select
                        REM **********************
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    For Each rows As DataRow In pl_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"
                                itemSQL = "insert into WMS_DO_PICKLIST_D " & _
                                        "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " & _
                                        "PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " & _
                                        "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values " & _
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " & _
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " & _
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " & _
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                '"('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.Text) & "', '" & nextNo & "', " & _
                        End Select
                        REM **********************
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    selectSql = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(nextNo) & "' "

                    plSeq = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSql, gConn, transaction), 0)

                    'For new added item, generate a empty picklist automatically
                    selectSql = "select DOD_ITM_CODE, DOD_PACK_KEY, DOD_PALLET_NO, DOD_QTY, DOD_BATCH_NO, DOD_QTY2, " & _
                                    "Convert(varchar,DOD_EXPIRY_DATE, " & gU.getConfig("DDFORMATNO") & ") as DOD_EXPIRY_DATE, " & _
                                    "Convert(varchar,DOD_MANU_DATE, " & gU.getConfig("DDFORMATNO") & ") as DOD_MANU_DATE " & _
                                "from WMS_DELV_ORDER_D d " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(nextNo) & "' " & _
                                "and not exists ( " & _
                                    "select 1 from WMS_DO_PICKLIST_D p " & _
                                    "where p.IMP_CODE = d.IMP_CODE " & _
                                    "and p.STORER_CODE = d.STORER_CODE " & _
                                    "and p.DO_CODE = d.DO_CODE " & _
                                    "and p.PLD_ITEM_NO = d.DOD_ITM_CODE " & _
                                    "and p.PLD_PACK_KEY = d.DOD_PACK_KEY) "

                    mplDt = gDB.getDataTable(selectSql, gConn, transaction)

                    'plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                    For i = 0 To mplDt.Rows.Count - 1
                        plSeq = plSeq + 1

                        If gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, "")) <> "" Then
                            itemSQL = "insert into WMS_DO_PICKLIST_D " & _
                                        "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_ITEM_QTY, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " & _
                                         "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, " & _
                                         "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                    "values " & _
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & _
                                         "'" & gU.dbEncode(CStr(plSeq)) & "', '" & Session("usr_id") & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim, "000")) & "', " & _
                                         "0, '', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("dod_qty").ToString.Trim, "0")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                         "0, " & _
                                         gU.convdbDate(gU.dbEncode(mplDt.Rows(i).Item("DOD_EXPIRY_DATE").ToString.Trim)) & ", " & _
                                         gU.convdbDate(gU.dbEncode(mplDt.Rows(i).Item("DOD_MANU_DATE").ToString.Trim)) & ", " & _
                                         "null, " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("dod_qty2").ToString.Trim, "0")) & "', " & _
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                            gDB.amendData(itemSQL, gConn, transaction)
                        End If
                    Next

                    Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode

                    sql_string = "update WMS_DELV_ORDER set " & _
                                    "DO_ISSUED_BY = '" & gU.dbEncode(DO_ISSUED_BY.Text.Trim) & "', " & _
                                    "DO_CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text.Trim) & "', " & _
                                    "DO_CUS_REF_NO = '" & gU.dbEncode(DO_CUS_REF_NO.Text.Trim) & "', " & _
                                    "DO_PROJECT_NO = '" & gU.dbEncode(DO_PROJECT_NO.SelectedValue) & "', " & _
                                    "DO_DATE = " & gU.convdbDate(gU.dbEncode(DO_DATE.Text.Trim)) & ", " & _
                                    "DO_TARGET_DELDATE = " & gU.convdbDate(gU.dbEncode(DO_TARGET_DELDATE.Text.Trim)) & ", " & _
                                    "DO_CONF_DELDATE = " & gU.convdbDate(gU.dbEncode(DO_CONF_DELDATE.Text.Trim)) & ", " & _
                                    "DO_CONF_DELTIME = "

                    If DO_CONF_DELDATE.Text.Trim <> "" AndAlso DO_CONF_DELTIME.Text.Trim <> "" Then
                        sql_string = sql_string & _
                            "to_date('" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', '" & gU.getConfig("DDFORMAT") & " hh24:mi'), "
                    Else
                        sql_string = sql_string & "null, "
                    End If

                    sql_string = sql_string & _
                                    "CUS_CODE = '" & gU.dbEncode(CUS_CODE.SelectedValue) & "', " & _
                                    "CUS_NAME = '" & gU.dbEncode(CUS_NAME.Text.Trim) & "', " & _
                                    "DO_ADDR1 = '" & gU.dbEncode(DO_ADDR1.Text.Trim) & "', " & _
                                    "DO_ADDR2 = '" & gU.dbEncode(DO_ADDR2.Text.Trim) & "', " & _
                                    "DO_ADDR3 = '" & gU.dbEncode(DO_ADDR3.Text.Trim) & "', " & _
                                    "DO_AREA_DEL = '" & gU.dbEncode(DO_AREA_DEL.Text.Trim) & "', " & _
                                    "DO_REGION_DEL = '" & gU.dbEncode(DO_REGION_DEL.Text.Trim) & "', " & _
                                    "DO_COUNTRY_DEL = '" & gU.dbEncode(DO_COUNTRY_DEL.Text.Trim) & "', " & _
                                    "DO_CUS_CONT = '" & gU.dbEncode(DO_CUS_CONT.Text.Trim) & "', " & _
                                    "DO_CUS_CONT_TEL = '" & gU.dbEncode(DO_CUS_CONT_TEL.Text.Trim) & "', " & _
                                    "DO_DRIVER = '" & gU.dbEncode(DO_DRIVER.Text.Trim) & "', " & _
                                    "DO_DRIVER_TEL = '" & gU.dbEncode(DO_DRIVER_TEL.Text.Trim) & "', " & _
                                    "DO_VEHICLE_NO = '" & gU.dbEncode(DO_VEHICLE_NO.Text.Trim) & "', " & _
                                    "DO_TOTL_PALLETS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_PALLETS.Text.Trim, "0")) & "', " & _
                                    "DO_TOTL_CARTONS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_CARTONS.Text.Trim, "0")) & "', " & _
                                    "DO_TOTL_BINS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_BINS.Text.Trim, "0")) & "', " & _
                                    "DO_REM = '" & gU.dbEncode(DO_REM.Text.Trim) & "', " & _
                                    "DO_TRACK_NO = '" & gU.dbEncode(DO_TRACK_NO.Text.Trim) & "', " & _
                                    "DO_FTRACK_NO = '" & gU.dbEncode(DO_FTRACK_NO.Text.Trim) & "', " & _
                                    "DO_CONSIGNEE = '" & gU.dbEncode(DO_CONSIGNEE.Text.Trim) & "', " & _
                                    "DO_CONSIGNEE_ADDR1 = '" & gU.dbEncode(DO_CONSIGNEE_ADDR1.Text.Trim) & "', " & _
                                    "DO_CONSIGNEE_ADDR2 = '" & gU.dbEncode(DO_CONSIGNEE_ADDR2.Text.Trim) & "', " & _
                                    "DO_CONSIGNEE_ADDR3 = '" & gU.dbEncode(DO_CONSIGNEE_ADDR3.Text.Trim) & "', " & _
                                    "DO_SHIP_TO = '" & gU.dbEncode(DO_SHIP_TO.Text.Trim) & "', " & _
                                    "DO_SHIP_ADDR1 = '" & gU.dbEncode(DO_SHIP_ADDR1.Text.Trim) & "', " & _
                                    "DO_SHIP_ADDR2 = '" & gU.dbEncode(DO_SHIP_ADDR2.Text.Trim) & "', " & _
                                    "DO_SHIP_ADDR3 = '" & gU.dbEncode(DO_SHIP_ADDR3.Text.Trim) & "', " & _
                                    "DO_PAY_TERMS = '" & gU.dbEncode(DO_PAY_TERMS.Text.Trim) & "', " & _
                                    "DO_TRADE_TERMS = '" & gU.dbEncode(DO_TRADE_TERMS.Text.Trim) & "', " & _
                                    "DO_SHIP_MODE = '" & gU.dbEncode(DO_SHIP_MODE.SelectedValue) & "', " & _
                                    "DO_INV_NO = '" & gU.dbEncode(DO_INV_NO.Text.Trim) & "', " & _
                                    "DO_PACK_LABEL_1 = '" & gU.dbEncode(DO_PACK_LABEL_1.Value.Trim) & "', " & _
                                    "DO_PACK_LABEL_2 = '" & gU.dbEncode(DO_PACK_LABEL_2.Value.Trim) & "', " & _
                                    "DO_PACK_LABEL_3 = '" & gU.dbEncode(DO_PACK_LABEL_3.Value.Trim) & "', " & _
                                    "DO_PACK_LABEL_QTY_1 = '" & gU.dbEncode(DO_PACK_LABEL_QTY_1.Value.Trim) & "', " & _
                                    "DO_PACK_LABEL_QTY_2 = '" & gU.dbEncode(DO_PACK_LABEL_QTY_2.Value.Trim) & "', " & _
                                    "DO_PACK_LABEL_QTY_3 = '" & gU.dbEncode(DO_PACK_LABEL_QTY_3.Value.Trim) & "', " & _
                                    "DO_TRANS_TYPE = " & "'" & gU.dbEncode(DO_TRANS_TYPE.SelectedValue) & "', " & _
                                    "DO_EDI_SIR_NO = " & "'" & gU.dbEncode(DO_EDI_SIR_NO.Text.Trim) & "', " & _
                                    "DO_TROLLEY_ID = " & "'" & gU.dbEncode(DO_TROLLEY_ID.Text.Trim) & "', " & _
                                    "DO_DRUM_ID = " & "'" & gU.dbEncode(DO_DRUM_ID.Text.Trim) & "', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                        uiFun.reOrderDetails(dt, "dod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag").ToString
                                Case "N"
                                    itemSQL = "insert into WMS_DELV_ORDER_D " & _
                                            "(IMP_CODE, STORER_CODE, DO_CODE, DOD_SEQ, DOD_DISP_SEQ, DOD_PALLET_NO, DOD_CARTON_NO, DOD_PACK_NO, DOD_ITM_CODE, DOD_PACK_KEY, " & _
                                            "DOD_ITM_DESC, DOD_PACK_TYPE, DOD_QTY, DOD_UOM, DOD_CUT_YN, DOD_QTY2, DOD_UOM2, DOD_PCS_UOM, DOD_TOTPCS, DOD_TOT_WGT, DOD_TOT_CBM, DOD_REM, DOD_TICKET_NO, DOD_VND_CODE, " & _
                                            "DOD_BATCH_NO, DOD_EXPIRY_DATE, DOD_MANU_DATE, " & _
                                             "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                            "values " & _
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_cut_yn").ToString.Trim, "N")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " & _
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " & _
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from WMS_DELV_ORDER_D " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                            "and DOD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update WMS_DELV_ORDER_D set " & _
                                                "DOD_DISP_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " & _
                                                "DOD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " & _
                                                "DOD_CARTON_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " & _
                                                "DOD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " & _
                                                "DOD_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " & _
                                                "DOD_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " & _
                                                "DOD_ITM_DESC = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " & _
                                                "DOD_PACK_TYPE = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " & _
                                                "DOD_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " & _
                                                "DOD_UOM = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " & _
                                                "DOD_CUT_YN = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_CUT_YN").ToString.Trim, "N")) & "', " & _
                                                "DOD_QTY2 = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " & _
                                                "DOD_UOM2 = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " & _
                                                "DOD_PCS_UOM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " & _
                                                "DOD_TOTPCS = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " & _
                                                "DOD_TOT_WGT = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " & _
                                                "DOD_TOT_CBM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " & _
                                                "DOD_REM = N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " & _
                                                "DOD_TICKET_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " & _
                                                "dod_vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " & _
                                                "DOD_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                                "DOD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                                "DOD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                            "and DOD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If


                    For Each rows As DataRow In pi_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"
                                itemSQL = "insert into wms_do_packing_d " & _
                                       "(imp_code,storer_code,do_code," & _
                                       "pad_pack_no,pad_pallet_no,pad_pack_type," & _
                                       "pad_totl_packs, pad_uom, pad_carton_no, " & _
                                       "pad_pack_key, pad_itm_code, pad_vnd_code, " & _
                                       "pad_batch_no, pad_qty, pad_length, " & _
                                       "pad_width, pad_height, pad_ref_no, " & _
                                       "pad_pack_by, pad_display_seq, pad_net_weight, pad_gross_weight, pad_min_packing, pad_origin,pad_pack_size,PAD_QTY_PER_CTN," & _
                                       "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                       "values " & _
                                       "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " & _
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " & _
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, "0")) & "', " & _
                                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "


                            Case "D", "R"
                                itemSQL = "delete from WMS_DO_PACKING_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                        "and PAD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "' "
                            Case Else
                                itemSQL = "update WMS_DO_PACKING_D set " & _
                                            "PAD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " & _
                                            "PAD_PACK_TYPE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " & _
                                            "PAD_TOTL_PACKS = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " & _
                                            "pad_uom = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " & _
                                           "pad_carton_no = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " & _
                                           "pad_pack_key = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " & _
                                           "pad_pack_size = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " & _
                                           "pad_itm_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " & _
                                           "pad_vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " & _
                                           "pad_batch_no = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, "")) & "', " & _
                                           "pad_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & ", " & _
                                           "pad_length = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & ", " & _
                                           "pad_width = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & ", " & _
                                           "pad_height = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & ", " & _
                                           "pad_ref_no = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " & _
                                           "pad_pack_by = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " & _
                                           "pad_display_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " & _
                                           "pad_net_weight = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " & _
                                           "pad_gross_weight = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " & _
                                           "pad_min_packing = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " & _
                                           "pad_origin = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " & _
                                           "PAD_QTY_PER_CTN = '" & gU.dbEncode(gU.decodeEmptyCInt(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, 0)) & "', " & _
                                           "sys_lub = '" & Session("usr_id") & "', " & _
                                            "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                            "and PAD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "'"
                        End Select
                        REM **********************

                        'Response.Write(itemSQL)
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    If Session("_M_OB_DO_TMP_pl_del_list") <> "" Then
                        Dim delSeqArray As String()

                        delSeqArray = Split(Session("_M_OB_DO_TMP_pl_del_list"), ", ")

                        For j = 0 To UBound(delSeqArray)
                            itemSQL = "delete from WMS_DO_PICKLIST_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                        "and PLD_SEQ = '" & gU.dbEncode(delSeqArray(j).Trim) & "' "

                            gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If

                    For Each rows As DataRow In pl_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"
                                itemSQL = "insert into WMS_DO_PICKLIST_D " & _
                                        "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " & _
                                        "PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " & _
                                        "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, " & _
                                         "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values " & _
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " & _
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " & _
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " & _
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " & _
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " & _
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            Case "D"
                                itemSQL = "delete from WMS_DO_PICKLIST_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                        "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "' "
                            Case Else
                                itemSQL = "update WMS_DO_PICKLIST_D set " & _
                                            "PLD_PICKED_BY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " & _
                                            "PLD_ITEM_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " & _
                                            "PLD_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " & _
                                            "PLD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " & _
                                            "PLD_ITEM_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " & _
                                            "PLD_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " & _
                                            "PLD_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " & _
                                            "PLD_FLOOR = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " & _
                                            "PLD_AREA = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " & _
                                            "PLD_RACK = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " & _
                                            "PLD_BIN = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " & _
                                            "PLD_IS_LOAN = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " & _
                                            "PLD_DO_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " & _
                                            "PLD_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, "")) & "', " & _
                                            "PLD_FOI_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " & _
                                            "PLD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " & _
                                            "PLD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " & _
                                            "PLD_SERIAL_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " & _
                                            "PLD_QTY2 = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " & _
                                            "sys_lub = '" & Session("usr_id") & "', " & _
                                            "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                        "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "'"

                        End Select
                        REM **********************

                        'Response.Write(itemSQL)
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    'Delete item not selected anymore
                    itemSQL = "delete from WMS_DO_PICKLIST_D " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                "and not exists ( " & _
                                    "select 1 from WMS_DELV_ORDER_D d " & _
                                    "where WMS_DO_PICKLIST_D.IMP_CODE = d.IMP_CODE " & _
                                    "and WMS_DO_PICKLIST_D.STORER_CODE = d.STORER_CODE " & _
                                    "and WMS_DO_PICKLIST_D.DO_CODE = d.DO_CODE " & _
                                    "and WMS_DO_PICKLIST_D.PLD_ITEM_NO = d.DOD_ITM_CODE " & _
                                    "and WMS_DO_PICKLIST_D.PLD_PACK_KEY = d.DOD_PACK_KEY) "

                    gDB.amendData(itemSQL, gConn, transaction)

                    'Delete item not selected anymore
                    itemSQL = "delete from WMS_DO_PICKLIST_D " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                "and ISNULL(PLD_ITEM_QTY, 0) = 0 " & _
                                "and not exists ( " & _
                                    "select 1 from WMS_DELV_ORDER_D d " & _
                                    "where WMS_DO_PICKLIST_D.IMP_CODE = d.IMP_CODE " & _
                                    "and WMS_DO_PICKLIST_D.STORER_CODE = d.STORER_CODE " & _
                                    "and WMS_DO_PICKLIST_D.DO_CODE = d.DO_CODE " & _
                                    "and WMS_DO_PICKLIST_D.PLD_ITEM_NO = d.DOD_ITM_CODE " & _
                                    "and WMS_DO_PICKLIST_D.PLD_PACK_KEY = d.DOD_PACK_KEY " & _
                                    "and ISNULL(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') = ISNULL(d.DOD_PALLET_NO, '000')) "

                    gDB.amendData(itemSQL, gConn, transaction)

                    selectSql = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

                    plSeq = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSql, gConn, transaction), 0)

                    'Comment the add new item SQL because sometimes the user would like to delete picklist -------------------------------------
                    ''For new added item, generate a empty picklist automatically
                    'selectSql = "select DOD_ITM_CODE, DOD_PACK_KEY, DOD_PALLET_NO, DOD_QTY, DOD_BATCH_NO " & _
                    '            "from WMS_DELV_ORDER_D d " & _
                    '            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    '            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    '            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                    '            "and not exists ( " & _
                    '                "select 1 from WMS_DO_PICKLIST_D p " & _
                    '                "where p.IMP_CODE = d.IMP_CODE " & _
                    '                "and p.STORER_CODE = d.STORER_CODE " & _
                    '                "and p.DO_CODE = d.DO_CODE " & _
                    '                "and p.PLD_ITEM_NO = d.DOD_ITM_CODE " & _
                    '                "and p.PLD_PACK_KEY = d.DOD_PACK_KEY) "

                    'mplDt = gDB.getDataTable(selectSql, gConn, transaction)

                    ''plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                    'For i = 0 To mplDt.Rows.Count - 1
                    '    plSeq = plSeq + 1

                    '    If gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, "")) <> "" Then
                    '        itemSQL = "insert into WMS_DO_PICKLIST_D " & _
                    '                    "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_ITEM_QTY, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " & _
                    '                     "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                    '                "values " & _
                    '                    "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " & _
                    '                     "'" & gU.dbEncode(CStr(plSeq)) & "', '" & Session("usr_id") & "', " & _
                    '                     "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim, "")) & "', " & _
                    '                     "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim, "")) & "', " & _
                    '                     "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim, "000")) & "', " & _
                    '                     "0, '', " & _
                    '                     "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("dod_qty").ToString.Trim, "0")) & "', " & _
                    '                     "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, "")) & "', " & _
                    '                     "0, " & _
                    '                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '        gDB.amendData(itemSQL, gConn, transaction)
                    '    End If
                    'Next

                    Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)


                If flag = "RELEASE" Then


                    updateSql = "update WMS_DELV_ORDER " & _
                                "set DO_STATUS = 'ASSIGNED' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSql, gConn, transaction)

                    updateSql = "update WMS_DO_PICKLIST_D " & _
                                "set PLD_STATUS = 'ASSIGNED' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSql, gConn, transaction)

                Else

                    If DO_STATUS.Text = "ASSIGNED" Then
                        updateSql = "update WMS_DO_PICKLIST_D " & _
                                    "set PLD_STATUS = 'ASSIGNED' " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                                    "and isnull(PLD_STATUS, 'NEW') = 'NEW' " & _
                                    "and exists (" & _
                                        "select 1 from WMS_DELV_ORDER d " & _
                                        "where d.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " & _
                                        "and d.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " & _
                                        "and d.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                        "and d.DO_STATUS = 'ASSIGNED') "

                        gDB.amendData(updateSql, gConn, transaction)
                    End If

                End If

                updatePLParentChild(gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    DO_CODE.Text = nextNo
                    DO_CODE_HF.Value = nextNo
                    Session("DO_CODE") = nextNo
                    Session("STORER_CODE") = STORER_CODE.SelectedValue
                    DO_CODE.ForeColor = Drawing.Color.Black
                    DO_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag = "" Then
                    reloadPage("Record has been saved successfully!")
                End If

                'Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "DO", DO_CODE.Text, "../../")

                'If flag <> "Y" AndAlso flag <> "P" Then
                '    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                '    Call BindGV()
                'End If

                save = True

            Catch ex As Exception
                'transaction.Rollback()
                'gConn.Close()
                'Response.Write(ex.Message)
                'uiFun.displayMsg(Me, "1008", "", Session("gLang"))

                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If

                If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                    Response.Write(ex.Message)
                    uiFun.displayMsg(Me, "1008", "", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                    uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
                End If

                save = False
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        Else
            save = False
        End If
    End Function

    Private Sub updatePLParentChild(ByRef gConn As SqlConnection, ByRef transaction As SqlTransaction)
        Dim updateSql As String

        updateSql = "update WMS_DO_PICKLIST_D " & _
                    "set PLD_SPLIT_FR = PLD_SEQ " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

        '"and PLD_SPLIT_FR is not null "

        gDB.amendData(updateSql, gConn, transaction)

        updateSql = "update WMS_DO_PICKLIST_D " & _
                    "set PLD_SPLIT_FR = (" & _
                        "select convert(varchar, max(convert(int, p.PLD_SEQ))) " & _
                        "from WMS_DO_PICKLIST_D p " & _
                        "where p.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " & _
                        "and p.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " & _
                        "and p.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                        "and p.PLD_ITEM_NO = WMS_DO_PICKLIST_D.PLD_ITEM_NO " & _
                        "and p.PLD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                        "and isnull(p.PLD_PALLET_NO, '000') = isnull(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') " & _
                        "and p.PLD_BATCH_NO = WMS_DO_PICKLIST_D.PLD_BATCH_NO " & _
                        "and convert(int, p.PLD_SEQ) < convert(int, WMS_DO_PICKLIST_D.PLD_SEQ)) " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " & _
                    "and exists (" & _
                        "select 1 from WMS_DO_PICKLIST_D p " & _
                        "where p.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " & _
                        "and p.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " & _
                        "and p.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                        "and p.PLD_ITEM_NO = WMS_DO_PICKLIST_D.PLD_ITEM_NO " & _
                        "and p.PLD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                        "and isnull(p.PLD_PALLET_NO, '000') = isnull(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') " & _
                        "and p.PLD_BATCH_NO = WMS_DO_PICKLIST_D.PLD_BATCH_NO " & _
                        "and convert(int, p.PLD_SEQ) < convert(int, WMS_DO_PICKLIST_D.PLD_SEQ)) "

        gDB.amendData(updateSql, gConn, transaction)

    End Sub


    Protected Sub BindGV()

        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim doCode, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If Session("DO_CODE") <> "" Then
            'pk_code = Session("GR_CODE")
            doCode = Session("DO_CODE")
            storerCode = Session("STORER_CODE")
        Else
            'pk_code = Request("GR_CODE")
            doCode = Server.UrlDecode(Request("DO_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        REM**********************
        REM Generate Dropdown List from WMS_COL_CODE Table
        ' uiFun.load_dropdownBy_ColCode(so_currency, "WMS_QUOTATION_HD.QUO_CURRENCY", Session("gLang"))
        REM **********************

        REM**********************
        REM Generate Dropdown List from WMS_COL_CODE Table
        'uiFun.load_dropdown(so_cus_code, "select cus_code, cus_name from wms_customer", "cus_code", "cus_code")
        REM **********************

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            DO_CODE.ForeColor = Drawing.Color.Red
            DO_STATUS.Text = "NEW"
            btnAttach.Visible = False
            'DO_DATE.Text = Format(Today(), "yyyy/MM/dd")
            DO_DATE.Text = Now.Date.ToString(gU.getConfig("DDFORMAT2"))
            IMP_CODE.Value = Session("IMP_CODE")
            REM **********************

            uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
                                          "storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                        "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                        "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

            uiFun.load_dropdown(DO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " & _
                                          "storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                          "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                          "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
        Else
            REM **********************
            REM Modify Here
            btnAttach.Visible = True

            REM Generate Data Table from Header
            SQLString = "SELECT IMP_CODE, STORER_CODE, DO_CODE, DO_STATUS, DO_ISSUED_BY, DO_CO_CODE, DO_CUS_REF_NO, DO_PROJECT_NO, " & _
                            "Convert(varchar, DO_DATE, " & gU.getConfig("DDFORMATNo") & ") as DO_DATE, " & _
                            "Convert(varchar,DO_TARGET_DELDATE, " & gU.getConfig("DDFORMATNo") & ") as DO_TARGET_DELDATE, " & _
                            "Convert(varchar,DO_CONF_DELDATE, " & gU.getConfig("DDFORMATNo") & ") as DO_CONF_DELDATE, " & _
                            "Convert(varchar(5),DO_CONF_DELTIME, 108) as DO_CONF_DELTIME, " & _
                            "CUS_CODE, CUS_NAME, DO_ADDR1, DO_ADDR2, DO_ADDR3, DO_AREA_DEL, " & _
                            "DO_REGION_DEL, DO_COUNTRY_DEL, DO_CUS_CONT, DO_CUS_CONT_TEL, DO_DRIVER, DO_DRIVER_TEL, DO_VEHICLE_NO, DO_TOTL_PALLETS, " & _
                            "DO_TOTL_CARTONS, DO_TOTL_BINS, DO_REM, DO_TRACK_NO, DO_FTRACK_NO, DO_POSTED_DATE, DO_POSTED_BY, DO_TEAM, DO_SCH_DATE, " & _
                            "DO_CONTRACT_TYPE, DO_AGREEMENT_NO, DO_INSTALLATION_REQ, DO_INITIAL_COLLECT, DO_INITIAL_COLLECT_AMT, DO_PAY_TERMS, " & _
                            "DO_DELIVERED_BY, DO_AREA_CODE, DO_DIS_CODE, DO_REGION_CODE, DO_DELI_STATUS, DO_TYPE, DO_ST_ID, DO_ST_SEQ, DO_BLOCK_DEL, " & _
                            "DO_FLOOR_DEL, DO_FLAT_DEL, DO_RTE_CODE, DO_SIGNATURE, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, DO_JOB_ID, DO_PDASTATUS, " & _
                            "DO_REASON_CODE, CUS_BILL_ADDR, CUST_CONT_NAME, CUST_CONT_PHONE, DO_ORG_ADDR, DO_ZONE, DO_SUB_ZONE, DO_ADDR_ID, " & _
                            "DO_ADDR_SEQ, DO_SUB_TYPE, DO_ENT_CODE, DO_ENT_DESC, DO_DELI_INSTR, DO_TARGET_AMPM, DO_RESCH_DELDATE, DO_CMPLT_DELDATE, " & _
                            "DO_COL_REMARK, DO_RCL_COD_AMT, DO_RCL_TONER_QTY, DO_RCL_CART_QTY, DO_RCL_PKG_QTY, DO_RCL_OTHER_QTY, DO_COL_COD_AMT, " & _
                            "DO_COL_TONER_QTY, DO_COL_CART_QTY, DO_COL_PKG_QTY, DO_COL_OTHER_QTY, DO_ISSUED_DATE, DO_APPR_DATE, DO_APPR_BY, " & _
                            "DO_CONF_DATE, DO_CONF_BY, DO_CANCEL_DATE, DO_CANCEL_BY, DO_DEL_DATE, DO_HOLD_DATE, DO_HOLD_BY, DO_RESCH_DATE, " & _
                            "DO_RESCH_BY, DO_COMPL_DATE, DO_COMPL_BY, DO_VERIFY_DATE, DO_VERIFY_BY, DLAL_ID, DO_LEADER, DO_PDA, DO_SALES_NAME, " & _
                            "DO_SALES_PHONE, DO_SALES_EMAIL, DO_MEM_TYPE, DO_DELI_TYPE, DO_PCK_ZONE, DO_PCK_SUB_ZONE, DO_PCK_ADDR_ID, " & _
                            "DO_PCK_ADDR_SEQ, DO_PCK_FLAT, DO_PCK_FLOOR, DO_PCK_BLOCK, DO_SCAN_CODE, DO_INSTALL_INSTR, DO_INSTALL_DATE, " & _
                            "DO_CODE_BF, DO_CODE_BF_TGT, DO_CODE_AF, DO_CODE_AF_TGT, DO_ORDER_NO, DO_RMA, DO_REF_DO_CODE, DO_SALES_CODE, DO_TARGET_DT_FR, " & _
                            "DO_TARGET_DT_TO, DO_LATEST_DELI_DATE, DO_MEM_DOC_TYPE, DO_REF_NO, DO_CTRL_ENT, DO_WHS_CODE, DO_SELF_PICK_YN, " & _
                            "DO_INSTALL_BY, DO_WT_TYPE, DO_MAP_ADDR, DO_REASON_DESC, DO_PDA_RS_CODE, DO_TERR_CODE, DO_ADDR_MAP_FLAG, DO_DAILY_SCH_FLAG, " & _
                            "DO_DAILY_SCH_CHG_YN, DO_RCV_CUTOFF_YN, DO_SUBMIT_BY, DO_SUBMIT_DATE, DO_APPROVER, DO_REJ_BY, DO_REJ_ON, DO_REJ_REASON, " & _
                            "DO_PCK_ADDR1, DO_MEMO_PRINT_YN, DO_MEMO_PRINT_DATE, DO_CONSIGNEE, DO_CONSIGNEE_ADDR1, DO_CONSIGNEE_ADDR2, " & _
                            "DO_CONSIGNEE_ADDR3, DO_SHIP_TO, DO_SHIP_ADDR1, DO_SHIP_ADDR2, DO_SHIP_ADDR3, DO_INV_NO, DO_TRADE_TERMS, " & _
                            "DO_SHIP_MODE, DO_PACK_LABEL_1, DO_PACK_LABEL_2, DO_PACK_LABEL_3, DO_PACK_LABEL_QTY_1, DO_PACK_LABEL_QTY_2, " & _
                            "DO_PACK_LABEL_QTY_3, DO_TRANS_TYPE, DO_CONFIRM_DATE, DO_DELI_FLAG, DO_EDI_SIR_NO, DO_TROLLEY_ID, DO_DRUM_ID " & _
                        "from WMS_DELV_ORDER " & _
                        "WHERE WMS_DELV_ORDER.DO_CODE = '" & gU.dbEncode(doCode) & "' " & _
                        "AND WMS_DELV_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                        "AND WMS_DELV_ORDER.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                DO_CODE.Text = dt.Rows(0).Item("DO_CODE").ToString
                DO_CODE_HF.Value = dt.Rows(0).Item("DO_CODE").ToString
                DO_STATUS.Text = dt.Rows(0).Item("DO_STATUS").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                DO_CUS_REF_NO.Text = dt.Rows(0).Item("DO_CUS_REF_NO").ToString
                'DO_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("DO_DATE").ToString)
                DO_DATE.Text = dt.Rows(0).Item("DO_DATE").ToString
                DO_ISSUED_BY.Text = dt.Rows(0).Item("DO_ISSUED_BY").ToString
                DO_CO_CODE.Text = dt.Rows(0).Item("DO_CO_CODE").ToString
                'DO_TARGET_DELDATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("DO_TARGET_DELDATE").ToString)
                'DO_CONF_DELDATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("DO_CONF_DELDATE").ToString)
                DO_TARGET_DELDATE.Text = dt.Rows(0).Item("DO_TARGET_DELDATE").ToString
                DO_CONF_DELDATE.Text = dt.Rows(0).Item("DO_CONF_DELDATE").ToString
                DO_CONF_DELTIME.Text = cU.chgToTime(dt.Rows(0).Item("DO_CONF_DELTIME").ToString)
                'DO_PROJECT_NO.SelectedValue = dt.Rows(0).Item("DO_PROJECT_NO").ToString
                'CUS_CODE.SelectedValue = dt.Rows(0).Item("CUS_CODE").ToString
                CUS_NAME.Text = dt.Rows(0).Item("CUS_NAME").ToString
                DO_ADDR1.Text = dt.Rows(0).Item("DO_ADDR1").ToString
                DO_ADDR2.Text = dt.Rows(0).Item("DO_ADDR2").ToString
                DO_ADDR3.Text = dt.Rows(0).Item("DO_ADDR3").ToString
                DO_AREA_DEL.Text = dt.Rows(0).Item("DO_AREA_DEL").ToString
                DO_REGION_DEL.Text = dt.Rows(0).Item("DO_REGION_DEL").ToString
                DO_COUNTRY_DEL.Text = dt.Rows(0).Item("DO_COUNTRY_DEL").ToString
                DO_CUS_CONT.Text = dt.Rows(0).Item("DO_CUS_CONT").ToString
                DO_CUS_CONT_TEL.Text = dt.Rows(0).Item("DO_CUS_CONT_TEL").ToString
                DO_DRIVER.Text = dt.Rows(0).Item("DO_DRIVER").ToString
                DO_DRIVER_TEL.Text = dt.Rows(0).Item("DO_DRIVER_TEL").ToString
                DO_VEHICLE_NO.Text = dt.Rows(0).Item("DO_VEHICLE_NO").ToString
                DO_TOTL_PALLETS.Text = dt.Rows(0).Item("DO_TOTL_PALLETS").ToString
                DO_TOTL_CARTONS.Text = dt.Rows(0).Item("DO_TOTL_CARTONS").ToString
                DO_TOTL_BINS.Text = dt.Rows(0).Item("DO_TOTL_BINS").ToString
                DO_REM.Text = dt.Rows(0).Item("DO_REM").ToString
                DO_TRACK_NO.Text = dt.Rows(0).Item("DO_TRACK_NO").ToString
                hide_DO_TRACK_NO.Value = dt.Rows(0).Item("DO_TRACK_NO").ToString
                DO_FTRACK_NO.Text = dt.Rows(0).Item("DO_FTRACK_NO").ToString
                'NEW FIELD

                DO_CONSIGNEE.Text = dt.Rows(0).Item("DO_CONSIGNEE").ToString
                DO_CONSIGNEE_ADDR1.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR1").ToString
                DO_CONSIGNEE_ADDR2.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR2").ToString
                DO_CONSIGNEE_ADDR3.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR3").ToString
                DO_SHIP_TO.Text = dt.Rows(0).Item("DO_SHIP_TO").ToString
                DO_SHIP_ADDR1.Text = dt.Rows(0).Item("DO_SHIP_ADDR1").ToString
                DO_SHIP_ADDR2.Text = dt.Rows(0).Item("DO_SHIP_ADDR2").ToString
                DO_SHIP_ADDR3.Text = dt.Rows(0).Item("DO_SHIP_ADDR3").ToString
                DO_TRADE_TERMS.Text = dt.Rows(0).Item("DO_TRADE_TERMS").ToString
                DO_PAY_TERMS.Text = dt.Rows(0).Item("DO_PAY_TERMS").ToString
                DO_SHIP_MODE.SelectedValue = dt.Rows(0).Item("DO_SHIP_MODE").ToString
                DO_INV_NO.Text = dt.Rows(0).Item("DO_INV_NO").ToString


                DO_PACK_LABEL_QTY_1.Value = dt.Rows(0).Item("DO_PACK_LABEL_QTY_1").ToString
                DO_PACK_LABEL_QTY_2.Value = dt.Rows(0).Item("DO_PACK_LABEL_QTY_2").ToString
                DO_PACK_LABEL_QTY_3.Value = dt.Rows(0).Item("DO_PACK_LABEL_QTY_3").ToString

                DO_PACK_LABEL_1.Value = dt.Rows(0).Item("DO_PACK_LABEL_1").ToString
                DO_PACK_LABEL_2.Value = dt.Rows(0).Item("DO_PACK_LABEL_2").ToString
                DO_PACK_LABEL_3.Value = dt.Rows(0).Item("DO_PACK_LABEL_3").ToString

                DO_EDI_SIR_NO.Text = dt.Rows(0).Item("DO_EDI_SIR_NO").ToString

                DO_TROLLEY_ID.Text = dt.Rows(0).Item("DO_TROLLEY_ID").ToString.Trim
                DO_DRUM_ID.Text = dt.Rows(0).Item("DO_DRUM_ID").ToString.Trim

                'END HERE
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                Select Case dt.Rows(0).Item("DO_STATUS").ToString
                    Case "NEW"
                        btnPost.Visible = False
                        btnUnPost.Visible = False
                        btnPick.Visible = True
                        btnUnPick.Visible = False
                    Case "PICKED"
                        btnPost.Visible = True
                        btnUnPost.Visible = False
                        btnPick.Visible = False
                        btnUnPick.Visible = True
                    Case "POSTED"
                        btnPost.Visible = False
                        btnUnPost.Visible = True
                        btnPick.Visible = False
                        btnUnPick.Visible = False
                    Case "CANCELLED"
                        btnPost.Visible = False
                        btnUnPost.Visible = False
                        btnPick.Visible = False
                        btnUnPick.Visible = False

                End Select


                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                If dt.Rows(0).Item("DO_PROJECT_NO").ToString = "" Then
                    uiFun.load_dropdown(DO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT  where " & _
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                        "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                        "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(DO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT WHERE " & _
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                          "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                          "and PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("DO_PROJECT_NO").ToString) & "' ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , , , True)
                End If


                lbl_PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("DO_PROJECT_NO").ToString) & "' ")
                CUS_NAME.Text = DB.getValueFromSQL("SELECT CUS_NAME FROM WMS_CUSTOMER WHERE CUS_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CUS_CODE").ToString) & "' ")

                uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
                                              "storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                              "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                              "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

                CUS_CODE.SelectedValue = dt.Rows(0).Item("CUS_CODE").ToString

                uiFun.load_dropdown(DO_TRANS_TYPE, "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'WMS_DELV_ORDER.DO_TRANS_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , , dt.Rows(0).Item("DO_TRANS_TYPE").ToString, True)

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here

        REM Generate Data Table from Detail 2

        SQLString = ""
        SQLString += "select pd.imp_code, pd.storer_code, pd.do_code, pd.pad_pack_no, "
        SQLString += "pd.pad_display_seq, pd.pad_pallet_no, pd.pad_pack_type, pd.pad_totl_packs, "
        SQLString += "pd.pad_uom, pd.pad_carton_no, pd.pad_pack_key, pd.pad_itm_code, pd.pad_vnd_code, "
        SQLString += "pd.pad_batch_no, pd.pad_qty, pd.pad_length, pd.pad_width, pd.pad_height, "
        SQLString += "pd.pad_ref_no, pd.pad_pack_by, pd.pad_net_weight, pd.pad_gross_weight, pd.pad_min_packing, "
        'SQLString += "ISNULL(d.aitm_pcs_per_pack, pd.pad_pack_key) as pad_pack_size, "
        SQLString += "pad_pack_size, pd.PAD_QTY_PER_CTN,"
        SQLString += "pd.pad_origin, pd.sys_cb, pd.sys_cd, pd.sys_lub, pd.sys_lud, 'U' as mFlag, 0 as itm_total_qty, " & _
                        "case when Charindex('-', pd.PAD_CARTON_NO )> 0 then " & _
                            "right('00000000000000000000' + replace(SUBSTRING(pd.PAD_CARTON_NO,1,CHARINDEX('-',pd.PAD_CARTON_NO)),'-',''),20) " & _
                        "else right('00000000000000000000' + ISNULL(pd.PAD_CARTON_NO,0),20) end as sort_col "
        SQLString += "from wms_do_packing_d pd "
        SQLString += "left outer join wms_item m on "
        SQLString += "pd.imp_code = m.imp_code and pd.storer_code = m.storer_code "
        SQLString += "and pd.pad_pack_key = m.pack_key "
        SQLString += "and pd.pad_itm_code = m.itm_code "
        SQLString += "left outer join wms_alt_vend_item d on m.imp_code = d.imp_code "
        SQLString += "and m.storer_code = d.storer_code and m.itm_code = d.itm_code "
        SQLString += "and m.pack_key = d.pack_key and pd.pad_vnd_code = d.vnd_code "
        SQLString += "where pd.do_code = '" & gU.dbEncode(doCode) & "' "
        SQLString += "and pd.storer_code = '" & gU.dbEncode(storerCode) & "' "
        SQLString += "and pd.imp_code = '" & Session("imp_code") & "' "

        SQLString += "order by sort_col, pd.pad_pack_no, pd.pad_itm_code, pd.pad_pack_key "
        'SQLString += " order by pd.pad_display_seq"
        REM **********************
        Session("_M_OB_DO_TMP_pi_dt") = gDB.getDataTable(SQLString)
        REM **********************

        REM Generate Data Table from Detail 3
        SQLString = "SELECT d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " & _
                        "d.PLD_WH, d.PLD_LOC, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, d.PLD_BATCH_NO, d.PLD_FOI_QTY, d.PLD_SERIAL_NO, d.PLD_QTY2, i.ITM_SERIAL_NO_YN, " & _
                        "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " & _
                        "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " & _
                        "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " & _
                        "Convert(varchar,ISNULL(d.pld_expiry_date, l.ILOC_EXPIRY_DATE), " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " & _
                        "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag " & _
                    "from WMS_DO_PICKLIST_D d " & _
                    "left outer join WMS_ITEM_LOC_BAL l " & _
                        "on d.STORER_CODE = l.STORER_CODE " & _
                        "and d.PLD_ITEM_NO = l.ITM_CODE " & _
                        "and d.PLD_PACK_KEY = l.PACK_KEY " & _
                        "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " & _
                        "and d.PLD_LOC = l.ILOC_LOC " & _
                        "and ISNULL(d.PLD_BATCH_NO, ' ') = ISNULL(l.ILOC_BATCH_NO, ' ') " & _
                    "left outer join WMS_DATE_CODE DC " & _
                        "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
                        "and d.IMP_CODE = DC.IMP_CODE " & _
                        "and d.STORER_CODE = DC.STORER_CODE " & _
                    "left outer join WMS_ITEM i " & _
                        "on d.IMP_CODE = i.IMP_CODE " & _
                        "and d.IMP_CODE = i.IMP_CODE " & _
                        "and d.PLD_ITEM_NO = i.ITM_CODE " & _
                        "and d.PLD_PACK_KEY = i.PACK_KEY " & _
                    "LEFT OUTER JOIN ( " & _
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, ' ') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " & _
                            "from wms_do_picklist_d p, wms_delv_order d " & _
                            "where d.imp_code = p.imp_code " & _
                            "and d.storer_code = p.storer_code " & _
                            "and d.do_code = p.do_code " & _
                            "and d.do_status = 'PICKED' " & _
                            "and d.do_code <> '" & gU.dbEncode(doCode) & "' " & _
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ' '), p.pld_loc) PICK_ITEM " & _
                        "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " & _
                        "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " & _
                        "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " & _
                        "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " & _
                        "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " & _
                        "AND ISNULL(d.PLD_BATCH_NO, ' ') = PICK_ITEM.PLD_BATCH_NO " & _
                        "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " & _
                    "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                    "and d.DO_CODE = '" & gU.dbEncode(doCode) & "' " & _
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "


        SQLString = SQLString & " order by  DC.DC_DATE_CODE, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, PLD_WH, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN "
        REM **********************
        Session("_M_OB_DO_TMP_pl_dt") = gDB.getDataTable(SQLString)
        REM **********************

        Dim objPlt As PickListTable
        objPlt = New PickListTable(DO_CO_CODE.Text, doCode, storerCode, Session("IMP_CODE"))

        objPlt.updateTotalQty(Session("_M_OB_DO_TMP_pl_dt"))

        objPlt.updateHoldBal(Session("_M_OB_DO_TMP_pl_dt"))

        objPlt = Nothing

        SQLString = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

        Session("_M_OB_DO_TMP_pl_seq") = DB.getValueFromSQL(SQLString)

        If Session("_M_OB_DO_TMP_pl_seq") = "" Then
            Session("_M_OB_DO_TMP_pl_seq") = "0"
        End If

        Session("_M_OB_DO_TMP_pl_del_list") = ""

        REM Generate Data Table from Detail
        SQLString = "SELECT d.DOD_SEQ, d.DOD_DISP_SEQ, d.DOD_PALLET_NO, d.DOD_CARTON_NO, d.DOD_PACK_NO, d.DOD_ITM_CODE, d.DOD_PACK_KEY, " & _
                        "d.DOD_ITM_DESC, d.DOD_PACK_TYPE, d.DOD_QTY, d.DOD_UOM, d.DOD_CUT_YN, d.DOD_QTY2, d.DOD_UOM2, d.DOD_PCS_UOM, d.DOD_TOTPCS, d.DOD_TOT_WGT, d.DOD_TOT_CBM, " & _
                        "d.DOD_REM, d.DOD_TICKET_NO, d.DOD_VND_CODE, d.DOD_BATCH_NO, d.IMP_CODE, d.STORER_CODE, d.DO_CODE, i.ITM_SKU_NO, i.ITM_DESC, " & _
                        "Convert(varchar,d.DOD_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_EXPIRY_DATE, " & _
                        "Convert(varchar,d.DOD_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_MANU_DATE, " & _
                        "'U' as mFlag, 1 as label_qty, '' as dod_ref_no, d.DOD_TROLLEY_ID, i.ITM_SERIAL_NO_YN " & _
                    "from WMS_DELV_ORDER_D d " & _
                    "Left outer join WMS_ITEM i " & _
                    "ON d.IMP_CODE = i.IMP_CODE " & _
                    "and d.STORER_CODE = i.STORER_CODE " & _
                    "and d.DOD_ITM_CODE = i.ITM_CODE " & _
                    "and d.DOD_PACK_KEY = i.PACK_KEY " & _
                    "where d.DO_CODE = '" & gU.dbEncode(doCode) & "' " & _
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " & _
                    "and d.IMP_CODE = '" & Session("IMP_CODE") & "' "


        SQLString = SQLString & " order by d.DOD_DISP_SEQ"
        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        Session("dt") = dt
        GridView1.DataBind()

    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_delv_order " & _
                     "set do_status = 'CANCELLED', " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and do_code = '" & gU.dbEncode(DO_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            'DO_STATUS.Text = "CANCELLED"

            'If save() Then
            '    ar.sec_write = "N"
            '    CancelBtn.Visible = False
            '    ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            '    uiFun.displayMsg(Me, "1011", "", Session("gLang"))
            'End If

            reloadPage("Record has been cancelled!")

        Catch ex As Exception
            'transaction.Rollback()
            'Response.Write(ex.Message)
            'uiFun.displayMsg(Me, "1008", "", Session("gLang"))

            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

            If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
            End If
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

    Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
        If save("RELEASE") Then
            reloadPage("Picking List has been released to the PDA successfully!")
        End If
    End Sub

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Dim updateSql, selectSql, SQLString As String
        Dim pl_dt As DataTable
        Dim checkdt As DataTable
        Dim itmKey As String
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim gConn As SqlConnection
        Dim transaction As SqlTransaction
        Dim lCOD_SEQ As String
        Dim successFlag As Boolean

        SQLString = "SELECT CO_STATUS FROM WMS_CUST_ORDER M " & _
            "WHERE CO_STATUS = 'CLOSED' " & _
            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The CO status is Closed, please re-open before post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The CO status is Closed, please re-open before post DO!", Session("gLang"))
            End If
            Exit Sub
        End If

        SQLString = "SELECT DO_STATUS FROM WMS_DELV_ORDER M " & _
                    "WHERE DO_STATUS = 'POSTED' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is Posted, please re-open before post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is Posted, please re-open before post DO!", Session("gLang"))
            End If
            Exit Sub
        End If

        If GridView1.Rows.Count > 0 Then
            'If validateAll("Y") Then
            If save("Y") Then
                If isValidPost() Then
                    gConn = gDB.getConnection()

                    transaction = gConn.BeginTransaction()

                    Try
                        qtyDict = New Dictionary(Of String, Double)
                        cbmDict = New Dictionary(Of String, Double)
                        wgtDict = New Dictionary(Of String, Double)

                        For Each rows As DataRow In dt.Rows
                            itmKey = gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("dod_batch_no").ToString.Trim, "")

                            If qtyDict.ContainsKey(itmKey) Then
                                qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0)
                                cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0), 14)
                                wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0), 14)
                            Else
                                qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0))
                                cbmDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0))
                                wgtDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0))
                            End If
                        Next

                        pl_dt = Session("_M_OB_DO_TMP_pl_dt")

                        Dim MANU_DATE As String = ""
                        Dim EXP_DATE As String = ""

                        'Sum total foi qty and picked qty for check stock bal validation
                        Dim objPlt As PickListTable
                        objPlt = New PickListTable(DO_CO_CODE.Text, DO_CODE.Text, STORER_CODE.SelectedValue, Session("IMP_CODE"))
                        objPlt.updateTotalQty(pl_dt)
                        objPlt.updateHoldBal(pl_dt)
                        objPlt = Nothing

                        'Check total bal (for hold stock, becoz hold stock hv no loc)
                        For Each rows As DataRow In pl_dt.Rows
                            If CDbl(rows.Item("hold_qty")) > 0 Then
                                If CDbl(rows.Item("total_item_qty")) > CDbl(rows.Item("avail_qty")) Then
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                          "don\'t have enough total stock to delivery!", Session("gLang"))
                                    Else
                                        uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                          "沒有足夠總貨存出貨!", Session("gLang"))
                                    End If
                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()
                                    Exit Sub
                                End If
                            End If
                        Next

                        'Check for each loc bal
                        For Each rows As DataRow In pl_dt.Rows
                            Dim SrchStr As String = ""
                            SrchStr += "select ILOC_BAL_QTY, Convert(varchar, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                            SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                            SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                            SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                            SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                            SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, ""), "000")) & "' "

                            If rows.Item("pld_batch_no").ToString.Trim <> "" Then
                                SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' "
                            Else
                                SrchStr += "AND ILOC_BATCH_NO is null "
                            End If

                            Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)


                            MANU_DATE = rows.Item("PLD_MANU_DATE").ToString.Trim
                            EXP_DATE = rows.Item("PLD_EXPIRY_DATE").ToString.Trim

                            If qtydt.Rows.Count > 0 Then
                                Dim stQTY As Integer = gU.decodeEmptyCInt(qtydt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
                                Dim plQTY As Integer = gU.decodeEmptyCInt(rows.Item("pld_item_qty").ToString.Trim, 0)

                                If stQTY < plQTY Then
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                          "don\'t have enough stock to delivery!", Session("gLang"))
                                    Else
                                        uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                          "沒有足夠貨存出貨!", Session("gLang"))
                                    End If

                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()
                                    Exit Sub
                                End If

                                If MANU_DATE = "" Then
                                    MANU_DATE = qtydt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim
                                End If
                                If EXP_DATE = "" Then
                                    EXP_DATE = qtydt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                                End If
                            Else
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                      "don\'t have stock to delivery!", Session("gLang"))
                                Else
                                    uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
                                                      "沒有此貨存出貨!", Session("gLang"))
                                End If

                                transaction.Rollback()
                                gConn.Close()
                                gConn.Dispose()
                                Exit Sub
                            End If

                            st.IO_SEQ = ""
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "000")
                            st.IO_CUST_CODE = CUS_CODE.SelectedValue
                            st.IO_WH = gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")
                            st.IO_AREA = gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")
                            st.IO_LOC = gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")
                            st.IO_DOC = "DO"
                            st.IO_DOC_ID = DO_CODE.Text.Trim
                            st.IO_QTY = gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "")
                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")
                            st.IO_MANU_DATE = MANU_DATE
                            st.IO_EXPIRY_DATE = EXP_DATE


                            'Dim tempUOM2, tempQty2 As String

                            'tempUOM2 = ""
                            'tempQty2 = ""

                            'For Each d_rows As DataRow In dt.Rows
                            '    If d_rows.Item("dod_itm_code").ToString.Trim = rows.Item("pld_item_no").ToString.Trim AndAlso d_rows.Item("dod_pack_key").ToString.Trim = rows.Item("pld_pack_key").ToString.Trim AndAlso _
                            '              d_rows.Item("dod_pallet_no").ToString.Trim = rows.Item("pld_pallet_no").ToString.Trim AndAlso d_rows.Item("dod_batch_no").ToString.Trim = rows.Item("pld_batch_no").ToString.Trim Then

                            '        tempUOM2 = d_rows.Item("DOD_UOM2").ToString.Trim
                            '        'tempQty2 = d_rows.Item("DOD_QTY2").ToString.Trim

                            '        Exit For
                            '    End If
                            'Next

                            st.IOS_QTY2 = rows.Item("pld_qty2").ToString.Trim

                            st.IOS_SERIAL_NO = rows.Item("pld_serial_no").ToString.Trim

                            st.setOrgSerialInfo(rows.Item("pld_serial_no").ToString.Trim, gConn, transaction)

                            'st.IOS_DRUM_LEVEL = "1"
                            'st.IOS_UOM2 = tempUOM2
                            'st.IOS_ORG_QTY2 = ""
                            'st.IOS_QTY2 = ""
                            'st.IOS_DRUM_ID = ""
                            'st.IOS_ORG_SERIAL_NO = ""
                            'st.IOS_SERIAL_NO = ""


                            itmKey = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" & _
                                     gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" & _
                                     gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" & _
                                     gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")

                            If qtyDict.ContainsKey(itmKey) Then
                                Dim tmpCbm As Decimal = 0
                                Dim tmpwgt As Decimal = 0
                                Dim tmpqty As Decimal = 0

                                tmpCbm = cbmDict.Item(itmKey)
                                tmpwgt = wgtDict.Item(itmKey)
                                tmpqty = qtyDict.Item(itmKey)

                                If tmpqty = 0 Then tmpqty = 1

                                st.IO_CBM = Math.Round(tmpCbm / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                st.IO_KG = Math.Round(tmpwgt / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                'st.IO_CBM = Math.Round(tmpCbm * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                'st.IO_KG = Math.Round(tmpwgt * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                            Else
                                st.IO_CBM = 0
                                st.IO_KG = 0
                            End If

                            If rows.Item("pld_serial_no").ToString.Trim = "" Then
                                updateSql = "update WMS_CUST_ORDER_D " & _
                                            "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
                                                "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                            "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                            "AND ISNULL(COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', ' ') " & _
                                            "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                            "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

                                gDB.amendData(updateSql, gConn, transaction)

                            Else
                                selectSql = "select COD_SEQ as value " & _
                                            "from WMS_CUST_ORDER_D " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                            "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                            "AND ISNULL(COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', ' ') " & _
                                            "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                            "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " & _
                                            "AND ISNULL(COD_POST_QTY, 0) < ISNULL(COD_QTY, 0) " & _
                                            "order by convert(int, COD_SEQ) "

                                lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

                                updateSql = "update WMS_CUST_ORDER_D " & _
                                            "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
                                                "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                            "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

                                gDB.amendData(updateSql, gConn, transaction)

                            End If

                            'Update coh_rel_qty for Hold stock logic
                            updateSql = "update wms_cust_order_hold " & _
                                        "set COH_REL_QTY = case when ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) < ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " then wms_cust_order_hold.coh_in_stock_qty else ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " & _
                                            "sys_lub = '" & Session("usr_id") & "', " & _
                                            "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                        "and wms_cust_order_hold.coh_status <> 'RELEASE' " & _
                                        "and ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) > 0 " & _
                                        "and exists (" & _
                                            "select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                                            "where wms_cust_order_hold.imp_code = d.imp_code " & _
                                            "and wms_cust_order_hold.storer_code = d.storer_code " & _
                                            "and wms_cust_order_hold.co_code = d.co_code " & _
                                            "and wms_cust_order_hold.cod_seq = d.cod_seq " & _
                                            "and c.imp_code = d.imp_code " & _
                                            "and c.storer_code = d.storer_code " & _
                                            "and c.co_code = d.co_code " & _
                                            "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                                            "and ISNULL(d.COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " & _
                                            "and ISNULL(d.COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', ' ') " & _
                                            "AND d.COD_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " & _
                                            "AND d.COD_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "') "

                            gDB.amendData(updateSql, gConn, transaction)

                            st.UpdateStockTrans("OUT", gConn, transaction)

                            st.UpdateStockBalTrans("OUT", gConn, transaction)

                            st.UpdateStockSerialTrans("OUT", gConn, transaction)

                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)

                        Next

                        updateSql = "update WMS_CUST_ORDER " & _
                                    "set CO_STATUS = 'CLOSED', " & _
                                        "sys_lub = '" & Session("usr_id") & "', " & _
                                        "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "AND NOT EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, ' ') " & _
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_CUST_ORDER " & _
                                    "set CO_STATUS = 'PARTIAL', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "AND EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
                                        "AND D.COD_POST_QTY > 0) " & _
                                    "AND EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, ' ') " & _
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_DELV_ORDER " & _
                                    "set DO_STATUS = 'POSTED', " & _
                                    "DO_POSTED_DATE = Getdate(), " & _
                                    "DO_POSTED_BY = '" & Session("usr_id") & "', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "

                        gDB.amendData(updateSql, gConn, transaction)

                        transaction.Commit()
                        gConn.Close()

                        'DO_STATUS.Text = "POSTED"
                        'ar.sec_viewMode = "Y"
                        'btnPost.Visible = False
                        'btnUnPost.Visible = True
                        'btnUnPick.Visible = False
                        'btnPick.Visible = False

                        'exceptionEditList.Add("saveConDate")
                        'exceptionEditList.Add("DO_CONF_DELTIME")
                        'exceptionEditList.Add("DO_CONF_DELDATE")
                        'exceptionEditList.Add("ImageButton3")

                        'ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

                        'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                        'saveBtn1.Visible = False
                        'saveBtn2.Visible = False
                        'saveConDate.Visible = True

                        'uiFun.displayMsg(Me, "1007", "", Session("gLang"))

                        'Call BindGV()

                        'selectItemBtn.Attributes("onclick") = "ItemLookUp('" & STORER_CODE.SelectedValue & "');"
                        'cSBBtn.Attributes("onclick") = "checkSB('" & STORER_CODE.SelectedValue & "');"
                        'cSBBtn.Enabled = True

                        successFlag = True

                        If gConn IsNot Nothing Then
                            If gConn.State = ConnectionState.Open Then
                                gConn.Close()
                                gConn.Dispose()
                            End If
                        End If

                    Catch ex As Exception
                        transaction.Rollback()
                        successFlag = False
                        gConn.Close()
                        gConn.Dispose()
                        Throw ex
                    End Try

                    If successFlag = True Then
                        reloadPage("Record has been Posted successfully!")

                        'Dim rmtPost As New RemotePost
                        'rmtPost.Url = "DOMain.aspx"
                        'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                        'rmtPost.Add("DO_CODE", DO_CODE.Text)
                        'rmtPost.alertMsg = "Record has been Posted successfully!"
                        'rmtPost.Post()
                    Else

                    End If
                End If
            End If
            'Else
            '    If Session("gLang") = "E" Then
            '        uiFun.displayMsgNew(updtPnlAlert, "", "Pick list location cannot be empty! Please check pick list.", Session("gLang"))
            '    Else
            '        uiFun.displayMsgNew(updtPnlAlert, "", "取貨位置不能空白! 請檢查取貨單", Session("gLang"))
            '    End If
            'End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "没有可供发布的物件!", Session("gLang"))
            End If
        End If
    End Sub

    Protected Sub addCOtoDO()
        Dim i, j As Integer

        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "
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

            Dim COItemList As String = ""
            Dim COListarray As String()
            Dim COSeqListarray As String()
            Dim allowGo As Boolean
            Dim qtyListArray As String()

            allowGo = True
            COListarray = Split(COList.Value, ", ")
            COSeqListarray = Split(COSeqList.Value, ", ")
            qtyListArray = Split(qtyList.Value, ", ")

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i = 0 To COListarray.Count - 1
                    qtyitemDict.Add(COListarray(i) & "_000_" & COSeqListarray(i), qtyListArray(i))
                Next
            End If

            If COListarray.Count = 0 Then
                If COList.Value <> "" Then
                    COItemList = COList.Value & "_000_" & COSeqList.Value
                    DO_CO_CODE.Text = COItemList
                End If
            Else
                For i = 0 To COListarray.Count - 1
                    COItemList = gU.appendToList(COItemList, COListarray(i) & "_000_" & COSeqListarray(i))
                    If DO_CO_CODE.Text <> "" Then
                        If DO_CO_CODE.Text <> COListarray(i) Then
                            allowGo = False
                        End If
                    End If
                    DO_CO_CODE.Text = COListarray(i)
                Next
            End If

            If allowGo = False Then
                DO_CO_CODE.Text = ""
                uiFun.displayMsgNew(updtPnlAlert.Page, "", "Your selected item related to differnet CO, please select again!", "")
            Else
                SQLString = "SELECT d.CO_CODE, d.COD_SEQ, d.COD_QTY, m.CO_REM, m.CO_SHIP_MODE, Convert(varchar,m.CO_TARGET_DELDATE," & DDFORMAT & ") as CO_TARGET_DELDATE, " & _
                                "d.COD_DISP_SEQ, d.COD_PALLET_NO, d.COD_ITM_CODE, d.COD_PACK_KEY, d.COD_ITM_DESC, d.COD_PACKING, " & _
                                "m.CO_TRACK_NO, d.COD_UOM, d.COD_PCS_UOM, d.COD_TOT_WGT, d.COD_TOT_CBM, d.COD_REM, d.COD_VND_CODE, d.COD_BATCH_NO, " & _
                                "m.CUS_CODE, m.CUS_NAME, m.CO_ADDR1, m.CO_ADDR2, m.CO_ADDR3, m.CO_AREA_DEL, m.CO_REGION_DEL, " & _
                                "m.CO_COUNTRY_DEL, m.CO_CUS_CONT, m.CO_CUS_CONT_TEL, m.CO_TRACK_NO, m.CO_INV_NO, m.CO_CUS_REF_NO, i.ITM_SKU_NO, i.ITM_DESC, " & _
                                "v.AITM_QTY_PER_CTN, v.AITM_VOL, v.CARTON_CBM, d.COD_TICKET_NO, " & _
                                "d.COD_UOM2, d.COD_QTY2, " & _
                                "Convert(varchar,d.COD_EXPIRY_DATE, " & DDFORMAT & ") as COD_EXPIRY_DATE, " & _
                                "Convert(varchar,d.COD_MANU_DATE, " & DDFORMAT & ") as COD_MANU_DATE " & _
                             "from  WMS_CUST_ORDER m " & _
                             "inner join WMS_CUST_ORDER_D d " & _
                             "ON d.CO_CODE = m.CO_CODE " & _
                             "and d.STORER_CODE = m.STORER_CODE " & _
                             "and d.IMP_CODE = m.IMP_CODE " & _
                             "left outer join WMS_ITEM i " & _
                             "on d.IMP_CODE = i.IMP_CODE " & _
                             "and d.STORER_CODE = i.STORER_CODE " & _
                             "and d.COD_ITM_CODE = i.ITM_CODE " & _
                             "and d.COD_PACK_KEY = i.PACK_KEY " & _
                             "left outer join V_ALT_VEND_ITEM v " & _
                             "on d.IMP_CODE = v.IMP_CODE " & _
                             "and d.STORER_CODE = v.STORER_CODE " & _
                             "and d.COD_ITM_CODE = v.ITM_CODE " & _
                             "and d.COD_PACK_KEY = v.PACK_KEY " & _
                             "where d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                             "and d.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                             "and d.CO_CODE + '_000_' + d.COD_SEQ IN ('" & Replace(COItemList, ", ", "', '") & "') "

                SQLString = SQLString & " order by d.COD_DISP_SEQ"
                REM **********************
                CO_dt = gDB.getDataTable(SQLString)

                If CO_dt.Rows.Count = 0 Then
                    Exit Sub
                End If

                For i = 0 To CO_dt.Rows.Count - 1
                    If Session("n_cur_seq") = "" Then
                        Session("n_cur_seq") = next_seq_no
                    Else
                        temp_seq_no = CInt(Session("n_cur_seq")) + 1
                        Session("n_cur_seq") = temp_seq_no.ToString
                    End If

                    Dim itmQty As Integer = 0

                    If qtyitemDict.ContainsKey(CO_dt.Rows(i).Item("CO_CODE").ToString & "_000_" & CO_dt.Rows(i).Item("COD_SEQ").ToString) Then
                        itmQty = gU.decodeEmptyCInt(qtyitemDict(CO_dt.Rows(i).Item("CO_CODE").ToString & "_000_" & CO_dt.Rows(i).Item("COD_SEQ").ToString), CInt(CO_dt.Rows(i).Item("COD_QTY")))
                    Else
                        Dim qtySQL As String = ""
                        Dim qtyDT As DataTable

                        qtySQL = " SELECT WMS_CUST_ORDER.CO_CODE, WMS_CUST_ORDER_D.COD_SEQ, " & _
                                 " WMS_CUST_ORDER_D.COD_QTY,  " & _
                                 " Case sign((ISNULL(WMS_CUST_ORDER_D.COD_QTY, 0) - ISNULL(delv.dod_qty,0))) when -1 then 0 else (ISNULL(WMS_CUST_ORDER_D.COD_QTY,0) - ISNULL(delv.dod_qty,0)) end as avai_qty " & _
                                 " from WMS_CUST_ORDER inner join WMS_CUST_ORDER_D " & _
                                 " on WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE " & _
                                 " and WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE " & _
                                 " and WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE " & _
                                 " left outer join ( " & _
                                 " 	select d.IMP_CODE, d.STORER_CODE, h.DO_CO_CODE, d.DOD_ITM_CODE,  " & _
                                 " 		d.DOD_PACK_KEY, d.DOD_PALLET_NO, d.DOD_BATCH_NO, sum(d.dod_qty) as dod_qty " & _
                                 " 	from WMS_DELV_ORDER h, WMS_DELV_ORDER_D d " & _
                                 " 	where h.IMP_CODE = d.IMP_CODE " & _
                                 " 	and h.STORER_CODE = d.STORER_CODE " & _
                                 " 	and h.DO_CODE = d.DO_CODE " & _
                                 " 	and ISNULL(h.DO_STATUS, ' ') <> 'CANCELLED' " & _
                                 " 	group by d.IMP_CODE, d.STORER_CODE, h.DO_CO_CODE, d.DOD_ITM_CODE, d.DOD_PACK_KEY, d.DOD_PALLET_NO, d.DOD_BATCH_NO) delv " & _
                                 " on WMS_CUST_ORDER_D.IMP_CODE = delv.IMP_CODE  " & _
                                 " and WMS_CUST_ORDER_D.STORER_CODE = delv.STORER_CODE  " & _
                                 " and WMS_CUST_ORDER_D.CO_CODE = delv.DO_CO_CODE  " & _
                                 " and WMS_CUST_ORDER_D.COD_ITM_CODE = delv.DOD_ITM_CODE  " & _
                                 " and WMS_CUST_ORDER_D.COD_PACK_KEY = delv.DOD_PACK_KEY  " & _
                                 " and ISNULL(WMS_CUST_ORDER_D.COD_PALLET_NO, '000') = ISNULL(delv.DOD_PALLET_NO , '000') " & _
                                 " and ISNULL(WMS_CUST_ORDER_D.COD_BATCH_NO, ' ') = ISNULL(delv.DOD_BATCH_NO , ' ') " & _
                                 " WHERE" & _
                                 " WMS_CUST_ORDER.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and WMS_CUST_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                                 " and WMS_CUST_ORDER.CO_CODE='" & gU.dbEncode(CO_dt.Rows(i).Item("CO_CODE").ToString) & "' AND WMS_CUST_ORDER_D.COD_SEQ='" & gU.dbEncode(CO_dt.Rows(i).Item("COD_SEQ").ToString) & "' " & _
                                 " and WMS_CUST_ORDER.CO_STATUS <> 'CLOSED' "


                        qtyDT = gDB.getDataTable(qtySQL)

                        If qtyDT.Rows.Count > 0 Then
                            itmQty = qtyDT.Rows(0).Item("avai_qty").ToString.Trim
                        Else
                            itmQty = CInt(CO_dt.Rows(i).Item("COD_QTY"))
                        End If

                    End If


                    dt.Rows.Add()

                    rows_count = dt.Rows.Count

                    REM **********************
                    REM Modify Here
                    dt.Rows(rows_count - 1).Item("DOD_SEQ") = Session("n_cur_seq").ToString
                    dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = CO_dt.Rows(i).Item("COD_DISP_SEQ")
                    dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = gU.decodeNullOrEmpty(CO_dt.Rows(i).Item("COD_PALLET_NO"), "000")
                    dt.Rows(rows_count - 1).Item("DOD_CARTON_NO") = ""
                    dt.Rows(rows_count - 1).Item("DOD_ITM_CODE") = CO_dt.Rows(i).Item("COD_ITM_CODE")
                    dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CO_dt.Rows(i).Item("ITM_SKU_NO")

                    dt.Rows(rows_count - 1).Item("DOD_PACK_KEY") = CO_dt.Rows(i).Item("COD_PACK_KEY")
                    dt.Rows(rows_count - 1).Item("DOD_ITM_DESC") = CO_dt.Rows(i).Item("COD_ITM_DESC")
                    dt.Rows(rows_count - 1).Item("ITM_DESC") = CO_dt.Rows(i).Item("ITM_DESC")
                    dt.Rows(rows_count - 1).Item("DOD_PACK_NO") = ""
                    dt.Rows(rows_count - 1).Item("DOD_PACK_TYPE") = CO_dt.Rows(i).Item("COD_PACKING")
                    dt.Rows(rows_count - 1).Item("DOD_TICKET_NO") = CO_dt.Rows(i).Item("COD_TICKET_NO")

                    'dt.Rows(rows_count - 1).Item("DOD_QTY") = CO_dt.Rows(i).Item("COD_QTY")
                    dt.Rows(rows_count - 1).Item("DOD_QTY") = itmQty

                    dt.Rows(rows_count - 1).Item("DOD_UOM") = CO_dt.Rows(i).Item("COD_UOM")
                    dt.Rows(rows_count - 1).Item("DOD_PCS_UOM") = CO_dt.Rows(i).Item("COD_PCS_UOM")

                    dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = itmQty * CInt(CO_dt.Rows(i).Item("COD_PCS_UOM"))

                    dt.Rows(rows_count - 1).Item("DOD_QTY2") = CO_dt.Rows(i).Item("COD_QTY2")
                    dt.Rows(rows_count - 1).Item("DOD_UOM2") = CO_dt.Rows(i).Item("COD_UOM2")

                    'dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = CO_dt.Rows(i).Item("COD_TOT_WGT")
                    'dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = CO_dt.Rows(i).Item("COD_TOT_CBM")

                    'Dim tmpNoOfCarton As Double

                    'If CO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString.Trim <> "" AndAlso CO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString.Trim <> "0" Then
                    '    tmpNoOfCarton = Math.Ceiling(itmQty / CDbl(CO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString.Trim))
                    'Else
                    '    tmpNoOfCarton = 0
                    'End If

                    'If tmpNoOfCarton <> 0 Then
                    '    dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("AITM_VOL"), "0"), 0) * tmpNoOfCarton
                    '    dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("CARTON_CBM"), "0"), 0) * tmpNoOfCarton
                    'Else
                    '    dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = CO_dt.Rows(i).Item("AITM_VOL")
                    '    dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = CO_dt.Rows(i).Item("CARTON_CBM")
                    'End If

                    If itmQty <> 0 Then
                        dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("AITM_VOL"), "0"), 0) * itmQty
                        dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("CARTON_CBM"), "0"), 0) * itmQty
                    Else
                        dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = CO_dt.Rows(i).Item("AITM_VOL")
                        dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = CO_dt.Rows(i).Item("CARTON_CBM")
                    End If

                    dt.Rows(rows_count - 1).Item("DOD_REM") = CO_dt.Rows(i).Item("COD_REM")

                    dt.Rows(rows_count - 1).Item("DOD_VND_CODE") = CO_dt.Rows(i).Item("COD_VND_CODE")
                    dt.Rows(rows_count - 1).Item("DOD_BATCH_NO") = CO_dt.Rows(i).Item("COD_BATCH_NO")

                    dt.Rows(rows_count - 1).Item("DOD_EXPIRY_DATE") = CO_dt.Rows(i).Item("COD_EXPIRY_DATE")
                    dt.Rows(rows_count - 1).Item("DOD_MANU_DATE") = CO_dt.Rows(i).Item("COD_MANU_DATE")

                    'dt.Rows(rows_count - 1).Item("DOD_NO_OF_CARTON") = 1


                    CUS_CODE.SelectedValue = CO_dt.Rows(i).Item("CUS_CODE").ToString
                    CUS_NAME.Text = CO_dt.Rows(i).Item("CUS_NAME").ToString
                    DO_ADDR1.Text = CO_dt.Rows(i).Item("CO_ADDR1").ToString
                    DO_ADDR2.Text = CO_dt.Rows(i).Item("CO_ADDR2").ToString
                    DO_ADDR3.Text = CO_dt.Rows(i).Item("CO_ADDR3").ToString
                    DO_AREA_DEL.Text = CO_dt.Rows(i).Item("CO_AREA_DEL").ToString
                    DO_REGION_DEL.Text = CO_dt.Rows(i).Item("CO_REGION_DEL").ToString
                    DO_COUNTRY_DEL.Text = CO_dt.Rows(i).Item("CO_COUNTRY_DEL").ToString
                    DO_CUS_CONT.Text = CO_dt.Rows(i).Item("CO_CUS_CONT").ToString
                    DO_CUS_CONT_TEL.Text = CO_dt.Rows(i).Item("CO_CUS_CONT_TEL").ToString
                    DO_TRACK_NO.Text = CO_dt.Rows(i).Item("CO_TRACK_NO").ToString
                    DO_SHIP_MODE.SelectedValue = CO_dt.Rows(i).Item("CO_SHIP_MODE").ToString
                    DO_TARGET_DELDATE.Text = CO_dt.Rows(i).Item("CO_TARGET_DELDATE").ToString
                    DO_INV_NO.Text = CO_dt.Rows(i).Item("CO_INV_NO").ToString
                    DO_REM.Text = CO_dt.Rows(i).Item("CO_REM").ToString
                    DO_CUS_REF_NO.Text = CO_dt.Rows(i).Item("CO_CUS_REF_NO").ToString
                    REM **********************
                    dt.Rows(rows_count - 1).Item("mFlag") = "N"
                Next
                dt.AcceptChanges()
                Session("dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()


                Dim pickListDt, pl_dt As DataTable
                Dim objPlt As PickListTable
                Dim plSeq As Integer

                objPlt = New PickListTable(DO_CO_CODE.Text, DO_CODE.Text, STORER_CODE.SelectedValue, Session("IMP_CODE"))

                pickListDt = objPlt.generatePickList(dt)

                objPlt = Nothing


                'Dim LOC_dt, pickListDt, pl_dt As DataTable
                ''Dim objPlt As PickListTable
                'Dim plSeq As Integer

                'SQLString = "SELECT l.ITM_CODE, l.PACK_KEY, l.ILOC_PALLET_NO, d.COD_QTY, l.ILOC_BAL_QTY, l.ILOC_LOC, " & _
                '                "l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO " & _
                '            "from (" & _
                '                "select COD_ITM_CODE, COD_PACK_KEY, ISNULL(COD_PALLET_NO, '000') as COD_PALLET_NO, sum(COD_QTY) as COD_QTY " & _
                '                "from WMS_CUST_ORDER_D " & _
                '                "where CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                '                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                '                "and IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                '                "and ISNULL(COD_QTY, 0) > 0 " & _
                '                "group by COD_ITM_CODE, COD_PACK_KEY, COD_PALLET_NO " & _
                '                 ") d inner join WMS_ITEM_LOC_BAL l on " & _
                '            "d.COD_ITM_CODE = l.ITM_CODE " & _
                '            "and d.COD_PACK_KEY = l.PACK_KEY " & _
                '            "and d.COD_PALLET_NO = l.ILOC_PALLET_NO " & _
                '            "left outer join WMS_DATE_CODE DC on " & _
                '            "l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
                '            "and l.IMP_CODE = DC.IMP_CODE " & _
                '            "and l.STORER_CODE = DC.STORER_CODE " & _
                '            "where l.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                '            "and l.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                '            "and ISNULL(l.ILOC_BAL_QTY, 0) > 0 " & _
                '            "order by l.ITM_CODE, DC.DC_CONV_DATE, l.PACK_KEY, l.ILOC_BATCH_NO, l.ILOC_BAL_QTY desc "

                'LOC_dt = gDB.getDataTable(SQLString)

                ''objPlt = New PickListTable(LOC_dt)

                ''pickListDt = objPlt.getPickList()

                ''objPlt = Nothing

                'pickListDt = New DataTable
                'pickListDt.Columns.Add("PLD_ITEM_NO", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_PACK_KEY", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_PALLET_NO", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_DO_QTY", Type.GetType("System.Double"))
                'pickListDt.Columns.Add("PLD_ITEM_QTY", Type.GetType("System.Double"))
                'pickListDt.Columns.Add("ILOC_BAL_QTY", Type.GetType("System.Double"))
                'pickListDt.Columns.Add("PLD_WH", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_LOC", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_FLOOR", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_AREA", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_RACK", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_BIN", Type.GetType("System.String"))
                'pickListDt.Columns.Add("PLD_BATCH_NO", Type.GetType("System.String"))

                'Dim currQty As Double = 0

                'For i = 0 To LOC_dt.Rows.Count - 1
                '    If gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0) > 0 AndAlso _
                '        LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim <> "" Then
                '        Dim itmK As String = LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
                '                                                 LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
                '                                                 LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString

                '        Dim doRow As DataRow() = dt.Select("DOD_ITM_CODE + '#_#' + DOD_PACK_KEY + '#_#' + DOD_PALLET_NO = '" & _
                '                                                        LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
                '                                                          LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
                '                                                          LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "'")
                '        Dim reqQty As Double = 0

                '        If doRow.Count > 0 Then
                '            reqQty = gU.decodeEmptyCdbl(doRow(0).Item("DOD_QTY"), 0)
                '        End If

                '        If currQty < reqQty Then
                '            Dim balQty As Double = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)

                '            Dim pNewRow = pickListDt.NewRow
                '            pNewRow.Item("PLD_ITEM_NO") = LOC_dt.Rows(i).Item("ITM_CODE").ToString.Trim
                '            pNewRow.Item("PLD_PACK_KEY") = LOC_dt.Rows(i).Item("PACK_KEY").ToString.Trim
                '            pNewRow.Item("PLD_PALLET_NO") = LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim

                '            pNewRow.Item("PLD_WH") = LOC_dt.Rows(i).Item("ILOC_WH").ToString.Trim
                '            pNewRow.Item("PLD_LOC") = LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                '            pNewRow.Item("PLD_FLOOR") = LOC_dt.Rows(i).Item("ILOC_FLOOR").ToString.Trim
                '            pNewRow.Item("PLD_AREA") = LOC_dt.Rows(i).Item("ILOC_AREA").ToString.Trim
                '            pNewRow.Item("PLD_RACK") = LOC_dt.Rows(i).Item("ILOC_RACK").ToString.Trim
                '            pNewRow.Item("PLD_BIN") = LOC_dt.Rows(i).Item("ILOC_BIN").ToString.Trim
                '            pNewRow.Item("PLD_BATCH_NO") = LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                '            pNewRow.Item("ILOC_BAL_QTY") = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)
                '            pNewRow.Item("PLD_DO_QTY") = reqQty

                '            If currQty + balQty >= reqQty Then
                '                pNewRow.Item("PLD_ITEM_QTY") = reqQty - currQty
                '                currQty = reqQty
                '            Else

                '                pNewRow.Item("PLD_ITEM_QTY") = balQty

                '                currQty += balQty

                '            End If

                '            pickListDt.Rows.Add(pNewRow)
                '            pickListDt.AcceptChanges()
                '        End If
                '    End If
                'Next

                pl_dt = Session("_M_OB_DO_TMP_pl_dt")
                plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                For i = 0 To pickListDt.Rows.Count - 1
                    pl_dt.Rows.Add()

                    plSeq = plSeq + 1

                    rows_count = pl_dt.Rows.Count

                    pl_dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                    pl_dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                    pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = pickListDt.Rows(i).Item("PLD_ITEM_NO")
                    pl_dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = pickListDt.Rows(i).Item("ITM_SKU_NO")
                    pl_dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = pickListDt.Rows(i).Item("PLD_PACK_KEY")
                    pl_dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = gU.decodeNullOrEmpty(pickListDt.Rows(i).Item("PLD_PALLET_NO"), "000")

                    pl_dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = pickListDt.Rows(i).Item("PLD_FOI_QTY")
                    pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                    pl_dt.Rows(rows_count - 1).Item("PLD_QTY2") = pickListDt.Rows(i).Item("PLD_QTY2")

                    pl_dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = pickListDt.Rows(i).Item("PLD_DO_QTY")
                    pl_dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = pickListDt.Rows(i).Item("ILOC_BAL_QTY")

                    pl_dt.Rows(rows_count - 1).Item("HOLD_QTY") = pickListDt.Rows(i).Item("HOLD_QTY")
                    pl_dt.Rows(rows_count - 1).Item("TOTAL_BAL") = pickListDt.Rows(i).Item("TOTAL_BAL")


                    pl_dt.Rows(rows_count - 1).Item("ILOC_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")

                    pl_dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")
                    pl_dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = pickListDt.Rows(i).Item("ILOC_MANU_DATE")

                    pl_dt.Rows(rows_count - 1).Item("PLD_WH") = pickListDt.Rows(i).Item("PLD_WH")
                    pl_dt.Rows(rows_count - 1).Item("PLD_LOC") = pickListDt.Rows(i).Item("PLD_LOC")
                    pl_dt.Rows(rows_count - 1).Item("PLD_FLOOR") = pickListDt.Rows(i).Item("PLD_FLOOR")
                    pl_dt.Rows(rows_count - 1).Item("PLD_AREA") = pickListDt.Rows(i).Item("PLD_AREA")
                    pl_dt.Rows(rows_count - 1).Item("PLD_RACK") = pickListDt.Rows(i).Item("PLD_RACK")
                    pl_dt.Rows(rows_count - 1).Item("PLD_BIN") = pickListDt.Rows(i).Item("PLD_BIN")
                    pl_dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                    pl_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                    pl_dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = pickListDt.Rows(i).Item("PLD_BATCH_NO")

                    pl_dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = pickListDt.Rows(i).Item("ITM_SERIAL_NO_YN")
                Next

                Dim isItemExist As Boolean

                'Check if any item is not existed in the pick list
                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i).Item("mFlag").ToString <> "D" Then
                        isItemExist = False
                        For j = 0 To pl_dt.Rows.Count - 1
                            If pl_dt.Rows(j).Item("mFlag").ToString <> "D" Then
                                If dt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim = pl_dt.Rows(j).Item("PLD_ITEM_NO").ToString.Trim AndAlso _
                                    dt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim = pl_dt.Rows(j).Item("PLD_PACK_KEY").ToString.Trim AndAlso _
                                    dt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim = pl_dt.Rows(j).Item("PLD_PALLET_NO").ToString.Trim AndAlso _
                                    (dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = "" OrElse dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = pl_dt.Rows(j).Item("PLD_BATCH_NO").ToString.Trim) Then

                                    isItemExist = True
                                    Exit For
                                End If
                            End If
                        Next

                        If Not isItemExist Then
                            pl_dt.Rows.Add()

                            plSeq = plSeq + 1

                            rows_count = pl_dt.Rows.Count

                            pl_dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                            pl_dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                            pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = dt.Rows(i).Item("DOD_ITM_CODE")
                            pl_dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = dt.Rows(i).Item("DOD_PACK_KEY")
                            pl_dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = gU.decodeNullOrEmpty(dt.Rows(i).Item("DOD_PALLET_NO"), "000")
                            pl_dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = 0
                            pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                            pl_dt.Rows(rows_count - 1).Item("PLD_QTY2") = dt.Rows(i).Item("DOD_QTY2")

                            pl_dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = dt.Rows(i).Item("DOD_QTY")

                            pl_dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = 0
                            pl_dt.Rows(rows_count - 1).Item("HOLD_QTY") = 0
                            pl_dt.Rows(rows_count - 1).Item("TOTAL_BAL") = 0

                            pl_dt.Rows(rows_count - 1).Item("PLD_WH") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_LOC") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_FLOOR") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_AREA") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_RACK") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_BIN") = ""
                            pl_dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                            pl_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                            pl_dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = dt.Rows(i).Item("DOD_BATCH_NO")

                            pl_dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = dt.Rows(i).Item("DOD_EXPIRY_DATE")
                            pl_dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = dt.Rows(i).Item("DOD_MANU_DATE")

                            pl_dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = dt.Rows(i).Item("ITM_SERIAL_NO_YN")

                        End If
                    End If
                Next

                Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)
                delPLZeroRow(pl_dt)

                Dim dtldt As DataTable = Session("dt")
                Dim pi_dt As DataTable = Session("_M_OB_DO_TMP_pi_dt")
                Dim nextSEQ As String = ""

                If dtldt IsNot Nothing AndAlso dtldt.Rows.Count > 0 Then

                    If pi_dt Is Nothing AndAlso pi_dt.Rows.Count = 0 Then
                        pi_dt = New DataTable
                        '" '"select pd.*, pd.pad_pack_key as pad_pack_size, pi.pld_item_qty, 'U' as mflag " & _
                        Dim SQLStr As String = "select pd.*, pi.pld_foi_qty, pi.pld_item_qty, 'U' as mflag " & _
                       " from wms_do_packing_d pd inner join WMS_DO_PICKLIST_D pi on " & _
                       "pd.imp_code = pi.imp_code " & _
                       "and pd.storer_code = pi.storer_code " & _
                       "and pd.pad_pack_key= pi.pld_pack_key " & _
                       "and pd.pad_itm_code = pi.pld_item_no " & _
                       "and pd.do_code = pi.do_code " & _
                        "and 1 = 0"

                        gDB.getDataTable(SQLStr, , , pi_dt)
                    Else
                        For Each piRow As DataRow In pi_dt.Rows
                            piRow.Item("mFlag") = "D"
                        Next

                        pi_dt.AcceptChanges()
                    End If

                    If nextSEQ = "" Then
                        Dim seqString As String = "select MAX(CAST(pad_pack_no AS int)) + 1 from wms_do_packing_d " & _
                                    "where IMP_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("IMP_CODE"))) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code").ToString)) & "' " & _
                                    "and DO_CODE = '" & gU.dbEncode(dtldt.Rows(0).Item("DO_CODE").ToString) & "' "

                        nextSEQ = DB.getValueFromSQL(seqString)

                        If nextSEQ Is Nothing OrElse nextSEQ = "" Then
                            nextSEQ = "1"
                        End If
                    End If

                    Dim indx As Integer = 1

                    For Each row As DataRow In pickListDt.Rows
                        Dim isPIItemExist As Boolean = False

                        If Not isPIItemExist Then
                            Dim newRow = pi_dt.NewRow

                            newRow.Item("pad_pack_no") = nextSEQ
                            newRow.Item("PAD_DISPLAY_SEQ") = indx

                            newRow.Item("pad_pallet_no") = gU.decodeNullOrEmpty(row.Item("pld_pallet_no"), "000")
                            newRow.Item("pad_itm_code") = row.Item("PLD_ITEM_NO")
                            newRow.Item("pad_batch_no") = row.Item("pld_batch_no")
                            newRow.Item("pad_pack_key") = row.Item("pld_pack_key")
                            'newRow.Item("pad_qty") = row.Item("pld_item_qty")
                            newRow.Item("pad_qty") = row.Item("pld_foi_qty")
                            'newRow.Item("pad_qty") = 0
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
                                                            "and m.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
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

                            pi_dt.Rows.Add(newRow)
                        End If
                    Next

                    pi_dt.AcceptChanges()
                End If

                Session("_M_OB_DO_TMP_pi_dt") = pi_dt

            End If
        End If
    End Sub

    Protected Sub addItemtoDO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "
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
            Dim DO_dt As DataTable

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


            SQLString = "SELECT M.*, D.* " & _
            "from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
            "where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            "and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
            "AND M.ITM_CODE + '_000_' + M.PACK_KEY IN ('" & Replace(itemPackList, ", ", "', '") & "') " & _
            "AND M.IMP_CODE *= D.IMP_CODE " & _
            "AND M.STORER_CODE *= D.STORER_CODE " & _
            "AND M.PACK_KEY *= D.PACK_KEY " & _
            "AND M.ITM_CODE *= D.ITM_CODE"

            SQLString = SQLString & " order by M.ITM_CODE, M.PACK_KEY, VND_CODE"
            REM **********************
            DO_dt = gDB.getDataTable(SQLString)
            For i As Integer = 0 To DO_dt.Rows.Count - 1
                If Session("n_cur_seq") = "" Then
                    Session("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(Session("n_cur_seq")) + 1
                    Session("n_cur_seq") = temp_seq_no.ToString
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("DOD_SEQ") = Session("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = Session("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("DOD_ITM_CODE") = DO_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = DO_dt.Rows(i).Item("ITM_SKU_NO")
                dt.Rows(rows_count - 1).Item("DOD_PACK_KEY") = DO_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = "000"
                dt.Rows(rows_count - 1).Item("DOD_ITM_DESC") = DO_dt.Rows(i).Item("ITM_NAME")
                dt.Rows(rows_count - 1).Item("ITM_DESC") = DO_dt.Rows(i).Item("ITM_DESC")
                dt.Rows(rows_count - 1).Item("DOD_QTY") = gU.decodeEmptyCInt(DO_dt.Rows(i).Item("ITM_BALANCE"), 0)
                dt.Rows(rows_count - 1).Item("DOD_UOM") = DO_dt.Rows(i).Item("AITM_UOM")
                dt.Rows(rows_count - 1).Item("DOD_PCS_UOM") = "1"
                dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = gU.decodeEmptyCInt(DO_dt.Rows(i).Item("ITM_BALANCE"), 0)
                dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = DO_dt.Rows(i).Item("ITM_BALANCE_KG")
                dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = DO_dt.Rows(i).Item("ITM_BALANCE_CBM")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
            Next
            dt.AcceptChanges()
            Session("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

        End If
    End Sub

    'This function should be same as PickList
    Private Sub delPLZeroRow(ByRef dt As DataTable)
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
                itmKey = dt.Rows(i).Item("pld_item_no").ToString.Trim & "#_#" & _
                        dt.Rows(i).Item("pld_pack_key").ToString.Trim & "#_#" & _
                        gU.decodeNullOrEmpty(dt.Rows(i).Item("pld_pallet_no").ToString.Trim, "000") & "#_#" & _
                        dt.Rows(i).Item("pld_batch_no").ToString.Trim

                'If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("pld_item_qty"), 0), 0)) > 0 Then
                If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("pld_foi_qty"), 0), 0)) > 0 Then
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
                If CInt(dt.Rows(delList(i)).Item("PLD_SEQ")) = CInt(Session("_M_OB_DO_TMP_pl_seq")) Then
                    Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) - 1)
                End If
                dt.Rows(delList(i)).Delete()
            Else
                dt.Rows(delList(i)).Item("mFlag") = "D"
            End If
        Next
        dt.AcceptChanges()
    End Sub

    Protected Sub DO_PROJECT_NO_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DO_PROJECT_NO.SelectedIndexChanged
        lbl_PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & DO_PROJECT_NO.SelectedValue & "' ")
    End Sub

    Protected Sub CUS_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_CODE.SelectedIndexChanged
        CUS_NAME.Text = DB.getValueFromSQL("select cus_name from wms_customer where storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and cus_code = '" & CUS_CODE.SelectedValue & "' ")

        Dim SQLSTR As String = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " & _
                        "cus_area_del, cus_region_del, cus_country_del, cus_cont_per_ord, cus_cont_tel_ord " & _
                        "from wms_customer where cus_code = '" & CUS_CODE.SelectedValue & "' " & _
                        "and storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                        "and imp_code = '" & Session("IMP_CODE") & "'"
        Dim cust_table As New DataTable

        cust_table = gDB.getDataTable(SQLSTR)

        If cust_table.Rows.Count > 0 Then
            CUS_NAME.Text = cust_table.Rows(0).Item("cus_name").ToString
            DO_ADDR1.Text = cust_table.Rows(0).Item("cus_addr1_del").ToString
            DO_ADDR2.Text = cust_table.Rows(0).Item("cus_addr2_del").ToString
            DO_ADDR3.Text = cust_table.Rows(0).Item("cus_addr3_del").ToString
            DO_AREA_DEL.Text = cust_table.Rows(0).Item("cus_area_del").ToString
            DO_REGION_DEL.Text = cust_table.Rows(0).Item("cus_region_del").ToString
            DO_COUNTRY_DEL.Text = cust_table.Rows(0).Item("cus_country_del").ToString()
            DO_CUS_CONT.Text = cust_table.Rows(0).Item("cus_cont_per_ord").ToString()
            DO_CUS_CONT_TEL.Text = cust_table.Rows(0).Item("cus_cont_tel_ord").ToString()
        Else
            CUS_NAME.Text = ""
            DO_ADDR1.Text = ""
            DO_ADDR2.Text = ""
            DO_ADDR3.Text = ""
            DO_AREA_DEL.Text = ""
            DO_REGION_DEL.Text = ""
            DO_COUNTRY_DEL.Text = ""
            DO_CUS_CONT.Text = ""
            DO_CUS_CONT_TEL.Text = ""
        End If
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        lbl_PRJ_NAME.Text = ""
        CUS_NAME.Text = ""
        DO_ADDR1.Text = ""
        DO_ADDR2.Text = ""
        DO_ADDR3.Text = ""
        DO_AREA_DEL.Text = ""
        DO_REGION_DEL.Text = ""
        DO_COUNTRY_DEL.Text = ""
        DO_CUS_CONT.Text = ""
        DO_CUS_CONT_TEL.Text = ""
        DO_CO_CODE.Text = ""

        uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
           " storer_code = '" & STORER_CODE.SelectedValue & "' " & _
         "and imp_code = '" & Session("IMP_CODE") & "' " & _
         "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

        uiFun.load_dropdown(DO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " & _
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                        "and imp_code = '" & Session("IMP_CODE") & "' " & _
                                        "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))

    End Sub

    Protected Sub btnUnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUnPost.Click
        Dim updateSql, selectSql, SQLString1 As String
        Dim gConn As SqlConnection
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim itmKey As String
        Dim pl_dt As DataTable
        Dim checkdt As DataTable
        Dim errorMsg As String = ""
        Dim successFlag As Boolean
        Dim lCOD_SEQ As String

        If Session.IsNewSession Then Exit Sub


        SQLString1 = "SELECT DO_STATUS FROM WMS_DELV_ORDER M " & _
                    "WHERE DO_STATUS <> 'POSTED' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString1)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is New, please post before un-post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is New, please post before un-post DO!", Session("gLang"))
            End If
            Exit Sub
        End If

        If GridView1.Rows.Count > 0 Then
            If DO_STATUS.Text = "POSTED" Then
                gConn = gDB.getConnection()

                Dim transaction As SqlTransaction
                transaction = gConn.BeginTransaction()

                Try
                    qtyDict = New Dictionary(Of String, Double)
                    cbmDict = New Dictionary(Of String, Double)
                    wgtDict = New Dictionary(Of String, Double)

                    For Each rows As DataRow In dt.Rows
                        itmKey = gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("dod_batch_no").ToString.Trim, "")

                        If qtyDict.ContainsKey(itmKey) Then
                            qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0)
                            cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0), 14)
                            wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0), 14)
                        Else
                            qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0))
                            cbmDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0), 14))
                            wgtDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0), 14))
                        End If
                    Next

                    pl_dt = Session("_M_OB_DO_TMP_pl_dt")

                    Dim ifPass As Boolean

                    Dim MANU_DATE As String = ""
                    Dim EXP_DATE As String = ""

                    For Each rows As DataRow In pl_dt.Rows

                        Dim SrchStr As String = ""
                        SrchStr += "select ILOC_BAL_QTY, Convert(varchar, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                        SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                        SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                        SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                        SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                        SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, ""), "000")) & "' "

                        If rows.Item("pld_batch_no").ToString.Trim <> "" Then
                            SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' "
                        Else
                            SrchStr += "AND ILOC_BATCH_NO is null "
                        End If

                        Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)

                        If qtydt.Rows.Count > 0 Then
                            MANU_DATE = qtydt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim
                            EXP_DATE = qtydt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                        End If

                        st.IO_SEQ = ""
                        st.STORER_CODE = STORER_CODE.SelectedValue
                        st.ITM_CODE = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")
                        st.PACK_KEY = gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")
                        st.PALLET_NO = gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "000")
                        st.IO_CUST_CODE = CUS_CODE.SelectedValue
                        st.IO_WH = gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")
                        st.IO_AREA = gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")
                        st.IO_LOC = gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")
                        st.IO_DOC = "DO"
                        st.IO_DOC_ID = DO_CODE.Text.Trim
                        st.IO_QTY = gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "")
                        st.lO_BATCH_NO = gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")
                        st.IO_MANU_DATE = MANU_DATE
                        st.IO_EXPIRY_DATE = EXP_DATE


                        st.IOS_QTY2 = rows.Item("pld_qty2").ToString.Trim

                        st.IOS_SERIAL_NO = rows.Item("pld_serial_no").ToString.Trim

                        st.setOrgSerialInfo(rows.Item("pld_serial_no").ToString.Trim, gConn, transaction)


                        itmKey = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" & _
                                 gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")

                        If qtyDict.ContainsKey(itmKey) Then
                            Dim tmpCbm As Decimal = 0
                            Dim tmpwgt As Decimal = 0
                            Dim tmpqty As Decimal = 0

                            tmpCbm = cbmDict.Item(itmKey)
                            tmpwgt = wgtDict.Item(itmKey)
                            tmpqty = qtyDict.Item(itmKey)

                            If tmpqty = 0 Then tmpqty = 1

                            st.IO_CBM = Math.Round(tmpCbm / tmpqty * CDbl(gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "0")), 2)
                            st.IO_KG = Math.Round(tmpwgt / tmpqty * CDbl(gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "0")), 2)
                            'st.IO_CBM = Math.Round(tmpCbm * CDbl(gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "0")), 2)
                            'st.IO_KG = Math.Round(tmpwgt * CDbl(gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "0")), 2)
                        Else
                            st.IO_CBM = 0
                            st.IO_KG = 0
                        End If


                        If rows.Item("pld_serial_no").ToString.Trim = "" Then
                            updateSql = "update WMS_CUST_ORDER_D " & _
                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
                                        "sys_lub = '" & Session("usr_id") & "', " & _
                                        "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                        "AND ISNULL(COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', ' ') " & _
                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                        Else
                            selectSql = "select COD_SEQ as value " & _
                                        "from WMS_CUST_ORDER_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                        "AND ISNULL(COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', ' ') " & _
                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " & _
                                        "AND ISNULL(COD_POST_QTY, 0) > " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & " " & _
                                        "order by convert(int, COD_SEQ) desc "

                            lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

                            updateSql = "update WMS_CUST_ORDER_D " & _
                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
                                            "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                            "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                        "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                        End If

                        'Update coh_rel_qty for Hold stock logic
                        updateSql = "update wms_cust_order_hold h " & _
                                    "set COH_REL_QTY = case when ISNULL(COH_REL_QTY,0) - " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " < 0 then 0 else ISNULL(COH_REL_QTY,0) - " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "and h.coh_status <> 'RELEASE' " & _
                                    "and ISNULL(h.coh_in_stock_qty, 0) > 0 " & _
                                    "and exists (select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                                    "where h.imp_code = d.imp_code " & _
                                    "and h.storer_code = d.storer_code " & _
                                    "and h.co_code = d.co_code " & _
                                    "and h.cod_seq = d.cod_seq " & _
                                    "and c.imp_code = d.imp_code " & _
                                    "and c.storer_code = d.storer_code " & _
                                    "and c.co_code = d.co_code " & _
                                    "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                                    "and ISNULL(d.COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " & _
                                    "and ISNULL(d.COD_BATCH_NO, ' ') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', ' ') " & _
                                    "AND d.COD_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " & _
                                    "AND d.COD_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "') "

                        gDB.amendData(updateSql, gConn, transaction)

                        ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKIN, gConn, transaction, errorMsg)

                        If Not ifPass Then
                            Exit For
                        End If
                    Next

                    If ifPass Then

                        updateSql = "update WMS_CUST_ORDER M " & _
                                    "set CO_STATUS = 'CLOSED', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "AND NOT EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = M.IMP_CODE " & _
                                        "AND D.STORER_CODE = M.STORER_CODE " & _
                                        "AND D.CO_CODE = M.CO_CODE " & _
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, ' ') " & _
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_CUST_ORDER M " & _
                                    "set CO_STATUS = 'PARTIAL', " & _
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "AND EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = M.IMP_CODE " & _
                                        "AND D.STORER_CODE = M.STORER_CODE " & _
                                        "AND D.CO_CODE = M.CO_CODE " & _
                                        "AND D.COD_POST_QTY > 0) " & _
                                    "AND EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = M.IMP_CODE " & _
                                        "AND D.STORER_CODE = M.STORER_CODE " & _
                                        "AND D.CO_CODE = M.CO_CODE " & _
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, ' ') " & _
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_CUST_ORDER M " & _
                                    "set CO_STATUS = 'NEW', " & _
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                                    "AND NOT EXISTS (" & _
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
                                        "WHERE D.IMP_CODE = M.IMP_CODE " & _
                                        "AND D.STORER_CODE = M.STORER_CODE " & _
                                        "AND D.CO_CODE = M.CO_CODE " & _
                                        "AND D.COD_POST_QTY > 0) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_DELV_ORDER " & _
                                    "set DO_STATUS = 'PICKED', " & _
                                        "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
                                        "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "

                        gDB.amendData(updateSql, gConn, transaction)

                        transaction.Commit()

                        'DO_STATUS.Text = "PICKED"
                        'ar.sec_viewMode = "N"
                        'ar.sec_write = "Y"
                        'btnPost.Visible = True
                        'btnUnPost.Visible = False

                        'ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

                        'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")
                        'cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);")
                        'cSBBtn.Enabled = True

                        'BindGV()

                        'Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "DO", DO_CODE.Text, "../../")

                        'uiFun.displayMsg(Me, "1013", "", Session("gLang"))
                        successFlag = True
                    Else
                        transaction.Rollback()
                        successFlag = False
                    End If


                    If gConn IsNot Nothing Then
                        If gConn.State = ConnectionState.Open Then
                            gConn.Close()
                            gConn.Dispose()
                        End If
                    End If

                Catch ex As Exception
                    transaction.Rollback()
                    successFlag = False
                    gConn.Close()

                    Throw ex
                End Try

                If successFlag = True Then
                    reloadPage("Record has been Posted successfully!")

                    'Dim rmtPost As New RemotePost
                    'rmtPost.Url = "DOMain.aspx"
                    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    'rmtPost.Add("DO_CODE", DO_CODE.Text)
                    'rmtPost.alertMsg = "Record has been unPosted successfully!"
                    'rmtPost.Post()
                Else
                    If errorMsg <> "" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
                    End If

                    'Dim rmtPost As New RemotePost
                    'rmtPost.Url = "DOMain.aspx"
                    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    'rmtPost.Add("DO_CODE", DO_CODE.Text)
                    'If errorMsg <> "" Then
                    '    rmtPost.alertMsg = "Un-Post Failed!\r\n" & errorMsg
                    'Else
                    '    rmtPost.alertMsg = "No item can be Un-Posted!"
                    'End If
                    'rmtPost.Post()
                End If

            End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "沒有可供取消發布的物件!", Session("gLang"))
            End If
        End If
    End Sub

    Protected Sub saveConDate_Click(sender As Object, e As System.EventArgs) Handles saveConDate.Click
        Dim add_sql As String = ""

        If DO_CONF_DELTIME.Text.Trim <> "" AndAlso Not gU.isTime(DO_CONF_DELTIME.Text.Trim) Then
            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid Delivery Time!", Session("gLang"))
            Exit Sub
        End If

        Dim updateSQL As String = ""
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction

        gconn = gDB.getConnection()
        transaction = gConn.BeginTransaction()

        Dim doCode, storerCode As String

        Try

            If Session("DO_CODE") <> "" Then
                'pk_code = Session("GR_CODE")
                doCode = Session("DO_CODE")
                storerCode = Session("STORER_CODE")
            Else
                'pk_code = Request("GR_CODE")
                doCode = Server.UrlDecode(Request("DO_CODE"))
                storerCode = Server.UrlDecode(Request("STORER_CODE"))
            End If

            If DO_CONF_DELDATE.Text.Trim <> "" AndAlso DO_CONF_DELTIME.Text.Trim <> "" Then
                add_sql = "Convert(datetime, '" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', 103) "
            Else
                add_sql = ""
            End If


            updateSQL = "Update wms_delv_order set " & _
                        "DO_CONF_DELDATE=convert(datetime, '" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & "'," & DDFORMAT & "), " & _
                        "DO_CONF_DELTIME=" & add_sql & " " & _
                        " Where imp_code='" & Session("imp_code") & "' AND Storer_code='" & gU.dbEncode(storerCode) & "' and do_code='" & gU.dbEncode(doCode) & "'"

            gDB.amendData(updateSQL, gconn, transaction)

            transaction.Commit()

            reloadPage("Record has been saved successfully!")

        Catch ex As Exception
            'transaction.Rollback()
            'gconn.Close()
            'Response.Write(ex.Message)
            'uiFun.displayMsg(Me, "1008", "", Session("gLang"))

            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

            If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
            End If

        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If
        End Try

        'uiFun.displayMsg(Me, "1007", "", Session("gLang"))
        'Call BindGV()

    End Sub

    Protected Sub btnPick_Click(sender As Object, e As System.EventArgs) Handles btnPick.Click
        Dim updateSQL As String = ""
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction

        If save("P") Then
            gconn = gDB.getConnection()
            transaction = gConn.BeginTransaction()

            Dim doCode, storerCode As String

            Try
                doCode = DO_CODE.Text
                storerCode = STORER_CODE.SelectedValue

                updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='PICKED', sys_LUD=Getdate(), sys_lub='" & gU.dbEncode(Session("usr_id")) & "' " & _
                            " Where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "' AND DO_CODE='" & gU.dbEncode(doCode) & "'"
                gDB.amendData(updateSQL)

                'Call BindGV()

                'btnPick.Visible = False
                'btnUnPick.Visible = True
                'btnPost.Visible = True

                'uiFun.displayMsgNew(updtPnlAlert, "", "The Order has been PICKED.", Session("gLang"))

                transaction.Commit()

                reloadPage("The order has been picked!")

            Catch ex As Exception
                'transaction.Rollback()
                'gconn.Close()
                'Response.Write(ex.Message)
                'uiFun.displayMsg(Me, "1008", "", Session("gLang"))

                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If

                If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                    Response.Write(ex.Message)
                    uiFun.displayMsg(Me, "1008", "", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                    uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
                End If

            Finally
                If gconn IsNot Nothing Then
                    If gconn.State = ConnectionState.Open Then
                        gconn.Close()
                        gconn.Dispose()
                    End If
                End If
            End Try
        End If
    End Sub

    Protected Sub btnUnPick_Click(sender As Object, e As System.EventArgs) Handles btnUnPick.Click
        Dim updateSQL As String = ""
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction

        gconn = gDB.getConnection()
        transaction = gConn.BeginTransaction()

        Dim doCode, storerCode As String

        Try
            doCode = DO_CODE.Text
            storerCode = STORER_CODE.SelectedValue

            updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='NEW', sys_LUD=Getdate(), sys_lub='" & gU.dbEncode(Session("usr_id")) & "' " & _
                        " Where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "' AND DO_CODE='" & gU.dbEncode(doCode) & "'"
            gDB.amendData(updateSQL)

            'Call BindGV()

            'btnPick.Visible = True
            'btnUnPick.Visible = False
            'btnPost.Visible = False

            'uiFun.displayMsg(Me, "", "The Order has been UNPICKED.", Session("gLang"))

            transaction.Commit()

            reloadPage("The order has been unpicked!")

        Catch ex As Exception
            'transaction.Rollback()
            'gconn.Close()
            'Response.Write(ex.Message)
            'uiFun.displayMsg(Me, "1008", "", Session("gLang"))

            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

            If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
            End If

        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Sub reloadPage(Optional ByVal alertMsg As String = "")
        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "RELOAD_PAGE", "reloadPage('" & gU.jsString(alertMsg) & "');", True)
    End Sub


End Class
