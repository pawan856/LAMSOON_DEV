Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

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

    Private gvCol() As String = {"DOD_DISP_SEQ", "DOD_WH_CODE", "DOD_TICKET_NO", "DOD_PALLET_NO", "DOD_CARTON_NO", "DOD_PACK_NO",
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
            lheader.Value = "Delivery Order Maintenance"
            lbl_DO_CODE.Text = "System DO Code:"
            lbl_DO_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_DO_CUS_REF_NO.Text = "Order No.:"
            lbl_DO_DATE.Text = "Date:"
            lbl_DO_ISSUED_BY.Text = "Issued By:"
            lbl_DO_CO_CODE.Text = "System CO No.:"
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
            lbl_DO_TO_STORER_REM.Text = "To Storer Remarks:"
            lbl_DO_STORER_REM.Text = "Storer Remarks:"
            'lbl_DO_TRACK_NO.Text = "Track No.:"
            lbl_DO_FTRACK_NO.Text = "Forwarder Tracking No.:"
            lbl_DO_TRANS_TYPE.Text = "Nature of Transaction"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            'New Field
            lbl_DO_CONSIGNEE.Text = "Consignee Name:"
            lbl_DO_CONSIGNEE_ADDR1.Text = "Consignee Address:"
            lbl_DO_POST_CODE.Text = "Post Code:"
            lbl_DO_DELIVERY_RMKS.Text = "Delivery Remarks:"
            lbl_DO_ARRIVAL_DATE.Text = "Returm Arrival Date:"
            lbl_DO_SHIP_TO.Text = "Ship to:"
            lbl_DO_SHIP_ADDR1.Text = "Ship Address:"
            lbl_DO_SHIP_MODE.Text = "Ship Mode:"
            lbl_DO_PAY_TERMS.Text = "Pay Terms:"
            lbl_DO_TRADE_TERMS.Text = "Trade Terms:"
            lbl_DO_INV_NO.Text = "Invoice No.:"
            lbl_DO_EDI_WIT_NO.Text = "SIR No.:"

            lbl_DO_PROVINCE.Text = "Providence"
            lbl_DO_CITY.Text = "City"

            lbl_DO_SENDER.Text = "Sender Name"
            lbl_DO_SENDER_COUNTRY.Text = "Sender Country"
            lbl_ROUTE_ID.Text = "ROUTE"

            lbl_DO_SENDER_PROVINCE.Text = "Sender Province"
            lbl_DO_SENDER_REGION.Text = "Sender Region"
            lbl_DO_SENDER_ADDR.Text = "Sender Address"
            lbl_DO_SENDER_TEL.Text = "Sender Telephone"

            btnNOTE.Text = "Notes"

            'End here
            saveStorerBtn.Text = "Save"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            'newrow.Text = "Add"
            'btnPost.Text = "Post"
            btnPostChk.Text = "Post"
            'selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveStorerBtn.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting) "");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting"")){getLoad();}else{return false;}"
            btnPostChk.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting"")){getLoad();}else{return false;}"
            btnUnPost.OnClientClick = "return confirm(""Are you sure to un-post this record?\r\n(Please save your work before Un-Posting)"");"
            If Session("pagemode") = "N" Then
                DO_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Value = "提貨單維護"
            lbl_DO_CODE.Text = "系統提貨號碼:"
            lbl_DO_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_DO_CUS_REF_NO.Text = "訂單號碼:"
            lbl_DO_DATE.Text = "日期:"
            lbl_DO_ISSUED_BY.Text = "核發者:"
            lbl_DO_CO_CODE.Text = "系統訂單號:"
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
            'lbl_DO_TRACK_NO.Text = "追查編號:"
            lbl_DO_FTRACK_NO.Text = "運送追查編號:"
            lbl_DO_TRANS_TYPE.Text = "貨單性質"
            lbl_DO_REM.Text = "備註:"
            lbl_DO_TO_STORER_REM.Text = "給貨主備註:"
            lbl_DO_STORER_REM.Text = "貨主備註:"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"

            REM New Field
            lbl_DO_CONSIGNEE.Text = "收貨人:"
            lbl_DO_CONSIGNEE_ADDR1.Text = "收貨地址:"
            lbl_DO_POST_CODE.Text = "郵政編號:"
            lbl_DO_DELIVERY_RMKS.Text = "收貨備註:"
            lbl_DO_ARRIVAL_DATE.Text = "回倉日期:"
            lbl_DO_SHIP_TO.Text = "運送到:"
            lbl_DO_SHIP_ADDR1.Text = "運送地址:"
            lbl_DO_SHIP_MODE.Text = "運送模式:"
            lbl_DO_PAY_TERMS.Text = "付款條款:"
            lbl_DO_TRADE_TERMS.Text = "貿易條款:"
            lbl_DO_INV_NO.Text = "發票號碼:"
            lbl_DO_EDI_WIT_NO.Text = "先生編號:"
            lbl_DO_TROLLEY_ID.Text = "手推車號"
            lbl_DO_DRUM_ID.Text = "鼓號"

            lbl_DO_PROVINCE.Text = "省"
            lbl_DO_CITY.Text = "市"

            lbl_DO_SENDER.Text = "寄件人姓名"
            lbl_DO_SENDER_COUNTRY.Text = "寄件人國家"
            lbl_ROUTE_ID.Text = "路線"
            lbl_DO_SENDER_PROVINCE.Text = "寄件人省份"
            lbl_DO_SENDER_REGION.Text = "寄件人地區"
            lbl_DO_SENDER_ADDR.Text = "寄件人地址"
            lbl_DO_SENDER_TEL.Text = "寄件人電話"
            btnNOTE.Text = "工單"

            ' New Field
            saveStorerBtn.Text = "保存"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            CancelBtn.Text = "取消"
            btnBack2.Value = "返回"
            btnBack.Value = "返回"
            'newrow.Text = "新增"
            btnRefresh.Value = "重新整理"
            btnPrtLbl.Text = "打印物品標籤"
            cSBBtn.Text = "檢查貨品存庫"
            btnPick.Text = "執貨"
            btnPost.Text = "出庫"
            btnPostChk.Text = "出庫"
            btnUnPost.Text = "反出庫"
            'btnRelease.Text = " 釋放"
            btnPrintPalletLabel.Text = "打印托盤標籤"
            'selectItemBtn.Text = "選擇物品"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveStorerBtn.OnClientClick = "return confirm(""確定保存資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定保存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定保存資料?"");"
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
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
        ROUTE_ID.CssClass = "REQUIRED"
        REM **********************

        'selectItemBtn.Visible = False

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

        'Set Access Right
        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            If Session("usr_pref_storer") <> STORER_CODE.Text Then
                Response.End()
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


        Call changeLabel()

        Dim colIdx_StartWith As Integer = 1

        ar.addColDef("DOD_DISP_SEQ", "dod_disp_seq", colIdx_StartWith)
        ar.addColDef("DOD_WH_CODE", "DOD_WH_CODE", colIdx_StartWith)
        ar.addColDef("DOD_TICKET_NO", "DOD_TICKET_NO", colIdx_StartWith)
        ar.addColDef("DOD_PALLET_NO", "dod_pallet_no", colIdx_StartWith)
        ar.addColDef("DOD_CARTON_NO", "dod_carton_no", colIdx_StartWith)
        ar.addColDef("DOD_PACK_NO", "dod_pack_no", colIdx_StartWith)
        ar.addColDef("DOD_BATCH_NO", "dod_batch_no", colIdx_StartWith)
        ar.addColDef("DOD_MANU_DATE", "dod_manu_date", colIdx_StartWith)
        ar.addColDef("DOD_VND_CODE", "dod_itm_code", colIdx_StartWith)
        ar.addColDef("DOD_ITM_CODE", "dod_itm_code", colIdx_StartWith)
        ar.addColDef("ITM_SKU_NO", "itm_sku_no", colIdx_StartWith)
        ar.addColDef("DOD_ITM_DESC", "dod_itm_desc", colIdx_StartWith)
        ar.addColDef("DOD_RETURN_QTY", "dod_return_qty", colIdx_StartWith)
        'ar.addColDef("DOD_CUT_YN", "dod_cut_yn", colIdx_StartWith)

        'ar.addColDef("ITM_DESC", "itm_desc", colIdx_StartWith)
        ar.addColDef("DOD_PACK_TYPE", "dod_pack_type", colIdx_StartWith)
        ar.addColDef("DOD_PACK_KEY", "dod_pack_key", colIdx_StartWith)

        ar.addColDef("DOD_EXPIRY_DATE", "dod_expiry_date", colIdx_StartWith)
        ar.addColDef("DOD_QTY", "dod_qty", colIdx_StartWith)
        ar.addColDef("DOD_UOM", "dod_uom", colIdx_StartWith)
        ar.addColDef("DOD_PCS_UOM", "dod_pcs_uom", colIdx_StartWith)
        ar.addColDef("DOD_QTY2", "dod_qty2", colIdx_StartWith)
        ar.addColDef("DOD_UOM2", "dod_uom2", colIdx_StartWith)

        ar.addColDef("DOD_TOTPCS", "dod_totpcs", colIdx_StartWith)
        ar.addColDef("DOD_TOT_WGT", "dod_tot_wgt", colIdx_StartWith)
        ar.addColDef("DOD_TOT_CBM", "dod_tot_cbm", colIdx_StartWith)
        ar.addColDef("DOD_REM", "dod_rem", colIdx_StartWith)
        ar.addColDef("DOD_PK_QTY", "DOD_PK_QTY", colIdx_StartWith)
        ar.addColDef("DOD_SS_QTY", "DOD_SS_QTY", colIdx_StartWith)
        ar.addColDef("DOD_FL_CODE", "dod_fl_code", colIdx_StartWith)
        ar.addColDef("DOD_MIN_PROD_DATE", "DOD_MIN_PROD_DATE", colIdx_StartWith)
        ar.addColDef("DOD_MIN_SHELF_LIFE", "DOD_MIN_SHELF_LIFE", colIdx_StartWith)
        ar.addColDef("DOD_LAST_LOT", "DOD_LAST_LOT", colIdx_StartWith)
        ar.addColDef("DOD_MAX_LOT", "DOD_MAX_LOT", colIdx_StartWith)
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
            selectCOBtn.Attributes.Add("onclick", "COLookUp('" & DO_CO_CODE.ClientID & "', '" & DO_CO_CODE.ClientID & "', document.myform." & STORER_CODE.ClientID & ".value);return false;")
        End If

        If DO_CO_CODE.Text <> "" Then
            selectCOBtn.Visible = False
        End If

        REM Select RO button
        'If DO_STATUS.Text = "POSTED" Or DO_STATUS.Text = "CANCELLED" Then
        '    selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
        '    cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');")
        '    cSBBtn.Enabled = True
        'Else
        '    selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")
        '    cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);")
        '    cSBBtn.Enabled = True
        'End If
        REM ****************************************************************

        If DO_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue
        ElseIf DO_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
            'btnPost.Visible = False
            btnPostChk.Visible = False
            btnUnPost.Visible = True

            btnPick.Visible = False
            btnUnPick.Visible = False

            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

            exceptionEditList.Add("saveConDate")
            If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
                exceptionEditList.Add("DO_STORER_REM")
            Else
                exceptionEditList.Add("DO_STORER_REM")
                exceptionEditList.Add("DO_CONF_DELTIME")
                exceptionEditList.Add("DO_CONF_DELDATE")
            End If
            exceptionEditList.Add("ImageButton3")
            exceptionEditList.Add("CalendarExtender3")
        ElseIf DO_STATUS.Text = "PICKED" Or DO_STATUS.Text = "SHORT-PICKED" Then
            'btnPost.Visible = True
            btnPostChk.Visible = True
            btnPick.Visible = False
            btnUnPick.Visible = True
            btnUnPost.Visible = False
        Else
            If Session("pagemode") = "N" Then btnPick.Visible = False Else btnPick.Visible = True
            btnUnPick.Visible = False
            'btnPost.Visible = False
            btnPostChk.Visible = False
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
            'Dim checkGP As String = "select usr_id from wms_user_group_alloc where grp_code = 'ADMIN_GP' and usr_id = '" & Session("usr_id") & "'"
            'Dim checkGPDt As DataTable
            'checkGPDt = gDB.getDataTable(checkGP)
            'If checkGPDt.Rows.Count > 0 Then
            '    'btnUnPost.Enabled = True
            '    btnUnPost.Enabled = ar.hasBtnRight("BT_DO_UNPOST")
            'Else
            '    btnUnPost.Enabled = False
            'End If
            btnUnPost.Enabled = ar.hasBtnRight("BT_DO_UNPOST")
            saveBtn1.Visible = False
            saveBtn2.Visible = False
            If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            Else
                saveConDate.Visible = True
                saveConDate.Enabled = True
                DO_STORER_REM.Enabled = True
            End If
        End If
        cSBBtn.Enabled = True

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OB_DO','" & Session("imp_code") & "||" & Session("STORER_CODE") & "||" & Session("DO_CODE") & "','N');")

        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            saveStorerBtn.Visible = True
            saveStorerBtn.Enabled = True
            DO_STORER_REM.Enabled = True
            saveBtn1.Visible = False
            saveBtn2.Visible = False
            btnRefresh.Visible = False
            'btnRelease.Visible = False
            btnPick.Visible = False
            btnUnPick.Visible = False
            btnPost.Visible = False
            btnPostChk.Visible = False
            btnUnPost.Visible = False
            btnPrtLbl.Visible = False
            btnPrintPalletLabel.Visible = False
        Else
            saveStorerBtn.Visible = False
        End If
        btnNOTE.Enabled = True
        If moduleAction.Value = "SAVEOK" Then
            save()
        End If

        If moduleAction.Value = "POSTOK" Then
            postChkOK(False)
        End If

        'ar.setFieldCustomize(Me, "OB_DO", STORER_CODE.SelectedValue, "WMS_DELV_ORDER")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "OB_DO", STORER_CODE.SelectedValue, "WMS_DELV_ORDER_D", gvCol)
        'End If
    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here
        'Call cU.newChangeGVLabel(GridView1, "", "")
        'Call cU.newChangeGVLabel(GridView1, "Seq.", "序號")
        'Call cU.newChangeGVLabel(GridView1, "WH", "倉庫")
        'Call cU.newChangeGVLabel(GridView1, "Floor", "地板")
        'Call cU.newChangeGVLabel(GridView1, "Location", "位置")
        'Call cU.newChangeGVLabel(GridView1, "DO No.", "不要")
        'Call cU.newChangeGVLabel(GridView1, "Pallet No.", "板號")
        'Call cU.newChangeGVLabel(GridView1, "Carton No.", "箱號")
        'Call cU.newChangeGVLabel(GridView1, "Pack No", "封裝碼")
        ''Call cU.newChangeGVLabel(GridView1, "Lot No.", "批號")
        'Call cU.newChangeGVLabel(GridView1, "Lot No.", "批號")
        'Call cU.newChangeGVLabel(GridView1, "Last Lot No.", "最后一批＃")
        'Call cU.newChangeGVLabel(GridView1, "Max Lot No.", "最大手数")
        'Call cU.newChangeGVLabel(GridView1, "Vendor Code", "供應商代碼")
        'Call cU.newChangeGVLabel(GridView1, "Item Code", "系統物件號碼")
        'Call cU.newChangeGVLabel(GridView1, "Stock No.", "物品號碼")
        'Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        'Call cU.newChangeGVLabel(GridView1, "Return Qty", "退貨數量")
        'Call cU.newChangeGVLabel(GridView1, "Cut", "剪")
        'Call cU.newChangeGVLabel(GridView1, "Pack Type", "封裝形式")
        'Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        'Call cU.newChangeGVLabel(GridView1, "Expiry Date", "失效日期")
        'Call cU.newChangeGVLabel(GridView1, "SS Qty", "SS数量")
        'Call cU.newChangeGVLabel(GridView1, "PK Qty", "PK数量")
        'Call cU.newChangeGVLabel(GridView1, "Qty", "數量")
        'Call cU.newChangeGVLabel(GridView1, "Min Prod Date", "最低生产日期")
        'Call cU.newChangeGVLabel(GridView1, "Min Shelf Life", "最低保质期")
        'Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        'Call cU.newChangeGVLabel(GridView1, "Number per UOM", "單位件數")
        'Call cU.newChangeGVLabel(GridView1, "Qty2", "數量2")
        'Call cU.newChangeGVLabel(GridView1, "UOM2", "單位2")
        'Call cU.newChangeGVLabel(GridView1, "Total Number", "總件數")
        'Call cU.newChangeGVLabel(GridView1, "Total Weight(kg)", "總重量")
        'Call cU.newChangeGVLabel(GridView1, "Total Volumne (cbm)", "總體積(cbm)")
        'Call cU.newChangeGVLabel(GridView1, "Remarks", "備註")
        'Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(saveBtn1)
        sm.RegisterAsyncPostBackControl(saveBtn2)
        sm.RegisterAsyncPostBackControl(saveStorerBtn)

        sm.RegisterAsyncPostBackControl(selectCOBtn)
        sm.RegisterAsyncPostBackControl(saveConDate)
        sm.RegisterAsyncPostBackControl(CancelBtn)
        sm.RegisterAsyncPostBackControl(cSBBtn)
        sm.RegisterAsyncPostBackControl(btnPick)
        sm.RegisterAsyncPostBackControl(btnUnPick)
        'sm.RegisterAsyncPostBackControl(btnPost)
        sm.RegisterAsyncPostBackControl(btnPostChk)
        sm.RegisterAsyncPostBackControl(btnUnPost)
        'sm.RegisterAsyncPostBackControl(btnRelease)

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

        If Request("READLOAD_JS") = "POST_CHK" Then
            INIT_JS.Value = "POST_CHK"
        Else
            INIT_JS.Value = ""
        End If
    End Sub

    Private Sub setGeneralControls()
        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
        Else
            If DO_CODE_HF.Value <> "" AndAlso DO_STATUS.Text = "New" Then
                'btnRelease.Visible = True
                'btnRelease.OnClientClick = "Return confirm('Confirm to release Picking List to PDA?');"
            End If
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

                'Dim strDate As String = If(DataBinder.Eval(e.Row.DataItem, "DOD_MIN_PROD_DATE").ToString.Trim IsNot DBNull.Value, DataBinder.Eval(e.Row.DataItem, "DOD_MIN_PROD_DATE").ToString.Trim, "")
                'If Not String.IsNullOrEmpty(strDate) Then
                '    Dim d As DateTime = Convert.ToDateTime(strDate)
                '    Dim reformatted As String = d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)
                '    CType(e.Row.FindControl("DOD_MIN_PROD_DATE"), TextBox).Text = reformatted
                'End If

                'strDate = If(DataBinder.Eval(e.Row.DataItem, "DOD_MIN_SHELF_LIFE").ToString.Trim IsNot DBNull.Value, DataBinder.Eval(e.Row.DataItem, "DOD_MIN_SHELF_LIFE").ToString.Trim, "")
                'If Not String.IsNullOrEmpty(strDate) Then
                '    Dim d As DateTime = Convert.ToDateTime(strDate)
                '    Dim reformatted As String = d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)
                '    CType(e.Row.FindControl("DOD_MIN_SHELF_LIFE"), TextBox).Text = reformatted
                'End If

                CType(e.Row.FindControl("DOD_MIN_PROD_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_MIN_PROD_DATE").ToString.Trim
                CType(e.Row.FindControl("DOD_MIN_SHELF_LIFE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_MIN_SHELF_LIFE").ToString.Trim

                CType(e.Row.FindControl("DOD_EXPIRY_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_EXPIRY_DATE").ToString.Trim
                'CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_MANU_DATE").ToString.Trim

                CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_EXPIRY_DATE").ToString.Trim
                If DataBinder.Eval(e.Row.DataItem, "DOD_EXPIRY_DATE").ToString.Trim = "" Then
                    CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text = ""
                Else
                    Dim d As DateTime = DateTime.ParseExact(CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text, "dd/MM/yyyy", CultureInfo.InvariantCulture)
                    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    CType(e.Row.FindControl("DOD_MANU_DATE"), Label).Text = reformatted
                End If

                CType(e.Row.FindControl("dod_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_QTY").ToString.Trim
                CType(e.Row.FindControl("DOD_PK_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_PK_QTY").ToString.Trim
                CType(e.Row.FindControl("DOD_SS_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_SS_QTY").ToString.Trim
                CType(e.Row.FindControl("DOD_LAST_LOT"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_LAST_LOT").ToString.Trim
                CType(e.Row.FindControl("DOD_MAX_LOT"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_MAX_LOT").ToString.Trim

                'Dim n1DropDown As DropDownList = CType(e.Row.FindControl("dod_uom"), DropDownList)
                'uiFun.load_dropdown(n1DropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'n1DropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim
                uiFun.load_dropdown(CType(e.Row.FindControl("dod_uom"), DropDownList), dtUOM, "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim)

                CType(e.Row.FindControl("dod_uom"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_uom").ToString.Trim

                CType(e.Row.FindControl("dod_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_QTY2").ToString.Trim

                CType(e.Row.FindControl("dod_return_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_RETURN_QTY").ToString.Trim

                'If DataBinder.Eval(e.Row.DataItem, "DOD_CUT_YN").ToString.Trim = "Y" Then
                '    CType(e.Row.FindControl("dod_cut_yn"), CheckBox).Checked = True
                'Else
                '    CType(e.Row.FindControl("dod_cut_yn"), CheckBox).Checked = False
                'End If

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
                CType(e.Row.FindControl("DOD_WH_CODE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_WH_CODE").ToString.Trim

                nDropDown = CType(e.Row.FindControl("dod_fl_code"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select distinct Convert(nvarchar(20),a.FL_NUM) as CODE,a.FL_NAME as NAME from WMS_WH_FL a where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "DOD_WH_CODE").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "dod_fl_code").ToString.Trim

                nDropDown = CType(e.Row.FindControl("DOD_LOC_WH"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "DOD_WH_CODE").ToString.Trim & "' and a.FL_NUM='" & DataBinder.Eval(e.Row.DataItem, "DOD_FL_CODE").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "DOD_LOC_WH").ToString.Trim

                'CType(e.Row.FindControl("DOD_LOC_WH"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DOD_LOC_WH").ToString.Trim
                CType(e.Row.FindControl("ITM_SERIAL_NO_YN"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim


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

    'Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
    '    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
    '        Dim rows_count As Integer = 0
    '        REM **********************
    '        REM Modify Here
    '        Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " & _
    '                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "
    '        REM **********************
    '        Dim nS_dt As New DataTable
    '        nS_dt = gDB.getDataTable(seq_string)
    '        Dim next_seq_no, temp_no As String
    '        Dim temp_seq_no As Integer = 1

    '        next_seq_no = ""

    '        If nS_dt.Rows.Count > 0 Then
    '            next_seq_no = nS_dt.Rows(0).Item(0).ToString()
    '        Else
    '            temp_no = "1"
    '        End If

    '        If next_seq_no = "" Then next_seq_no = "1"

    '        If Session("n_cur_seq") = "" Then
    '            Session("n_cur_seq") = next_seq_no
    '        Else
    '            temp_seq_no = CInt(Session("n_cur_seq")) + 1
    '            Session("n_cur_seq") = temp_seq_no.ToString
    '        End If


    '        dt.Rows.Add()
    '        rows_count = dt.Rows.Count

    '        REM **********************
    '        REM Modify Here
    '        dt.Rows(rows_count - 1).Item("dod_seq") = Session("n_cur_seq").ToString
    '        REM **********************
    '        dt.Rows(rows_count - 1).Item("mFlag") = "N"
    '        dt.AcceptChanges()

    '        Session("dt") = dt
    '        GridView1.DataSource = dt
    '        GridView1.DataBind()

    '    End If
    'End Sub

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

        If ROUTE_ID.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(Me, "", lbl_ROUTE_ID.Text & " cannot be empty! Please enter the value of ROUTE!", Session("gLang"))
            Else
                uiFun.displayMsgNew(Me, "", lbl_ROUTE_ID.Text & "不能空白!", Session("gLang"))
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

        If flag = "P" Then
            If DO_TROLLEY_ID.Text.Trim = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_TROLLEY_ID.Text & " cannot be empty!", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", lbl_DO_TROLLEY_ID.Text & "不能空白!", Session("gLang"))
                End If
                Return False
            End If
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
        If DO_TO_STORER_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Too many characters, maximum length of " & lbl_DO_TO_STORER_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "字數太多, " & lbl_DO_TO_STORER_REM.Text & "最多只限200字!", Session("gLang"))
            End If
            Return False
        End If
        If DO_STORER_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Too many characters, maximum length of " & lbl_DO_STORER_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "字數太多, " & lbl_DO_STORER_REM.Text & "最多只限200字!", Session("gLang"))
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

                If CType(GridView1.Rows(i).FindControl("itm_serial_no_yn"), HiddenField).Value = "Y" Then
                    If dtlQtyDict.ContainsKey(itmKey) Then
                        dtlQtyDict.Item(itmKey) = dtlQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty2"), TextBox).Text, 0)) * CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0))
                    Else
                        dtlQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty2"), TextBox).Text, 0)) * CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0)))
                    End If
                Else
                    If dtlQtyDict.ContainsKey(itmKey) Then
                        dtlQtyDict.Item(itmKey) = dtlQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0))
                    Else
                        dtlQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("dod_qty"), TextBox).Text, 0)))
                    End If
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

                If pl_dt.Rows(i).Item("itm_serial_no_yn").ToString.Trim = "Y" Then
                    If plQtyDict.ContainsKey(itmKey) Then
                        plQtyDict.Item(itmKey) = plQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_qty2"), 0), 0))
                    Else
                        plQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_qty2"), 0), 0)))
                    End If
                Else
                    If plQtyDict.ContainsKey(itmKey) Then
                        plQtyDict.Item(itmKey) = plQtyDict.Item(itmKey) + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_item_qty"), 0), 0))
                    Else
                        plQtyDict.Add(itmKey, CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(pl_dt.Rows(i).Item("pld_item_qty"), 0), 0)))
                    End If
                End If
            End If
        Next

        If postFlag = "Y" Then
            For i = 0 To pl_dt.Rows.Count - 1
                If pl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
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
                End If
            Next
        End If

        If postFlag = "Y" Then
            keys = dtlQtyDict.Keys

            For i = 0 To keys.Count - 1
                If plQtyDict.ContainsKey(keys(i)) Then
                    If Math.Round(dtlQtyDict.Item(keys(i)), 4) < Math.Round(plQtyDict.Item(keys(i)), 4) Then
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
                        'If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                        '    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                        'End If

                        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "CONFIRM_WARN", javaStr, True)

                        Return False
                    End If
                Else
                    'itmArray = Split(keys(i), "#_#")
                    'If Session("gLang") = "E" Then
                    '    uiFun.displayMsgNew(updtPnlAlert, "", "There is no Picking List for item, " & itmArray(0) & "-" & itmArray(1) & ", please check the Picking List!", Session("gLang"))
                    'Else
                    '    uiFun.displayMsgNew(updtPnlAlert, "", "在取貨單中找不到貨品, " & itmArray(0) & "-" & itmArray(1) & ", 請檢楂取貨單!", Session("gLang"))
                    'End If

                    'Return False
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

        selectSql = "select delv.imp_code, " &
                        "delv.storer_code, " &
                        "delv.pld_item_no as itm_code, " &
                        "i.itm_name, " &
                        "delv.pld_pack_key as pack_key, " &
                        "delv.pld_pallet_no as pallet_no, " &
                        "delv.pld_batch_no as batch_no, " &
                        "delv.pld_item_qty, " &
                        "case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end as hold_qty, " &
                        "ISNULL(case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) as itm_balance, " &
                        "ISNULL(case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) - " &
                            "ISNULL(case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end), " &
                            "0) as avail_qty " &
                    "from ( " &
                        "select imp_code, " &
                            "storer_code, " &
                            "pld_item_no, " &
                            "pld_pack_key, " &
                            "ISNULL(pld_pallet_no, '000') as pld_pallet_no, " &
                            "pld_batch_no, " &
                            "sum(pld_item_qty) as pld_item_qty " &
                        "from wms_do_picklist_d d " &
                        "where imp_code = " & cmdPa.AP(Session("IMP_CODE")) & " " &
                        "and storer_code = " & cmdPa.AP(STORER_CODE.SelectedValue) & " " &
                        "and DO_CODE = " & cmdPa.AP(DO_CODE.Text.Trim) & " " &
                        "group by imp_code, storer_code, pld_item_no, pld_pack_key, ISNULL(pld_pallet_no, '000'), pld_batch_no) delv, wms_item i, " &
                        "( " &
                        "select h.imp_code, " &
                            "h.storer_code, " &
                            "h.COH_ITM_CODE, " &
                            "h.COH_PACK_KEY, " &
                            "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                            "h.COH_BATCH_NO, " &
                            "sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
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
                        "and c.co_code <> " & cmdPa.AP(DO_CO_CODE.Text.Trim) & " " &
                        "group by h.imp_code, h.storer_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), h.COH_BATCH_NO) hold, " &
                        "( " &
                        "select h.imp_code, " &
                            "h.storer_code, " &
                            "h.COH_ITM_CODE, " &
                            "h.COH_PACK_KEY, " &
                            "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                            "sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
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
                        "and c.co_code <> " & cmdPa.AP(DO_CO_CODE.Text.Trim) & " " &
                        "group by h.imp_code, h.storer_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000')) hold2, " &
                        "( " &
                        "select l.imp_code, " &
                            "l.storer_code, " &
                            "l.ITM_CODE, " &
                            "l.PACK_KEY, " &
                            "ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                            "l.ILOC_BATCH_NO, " &
                            "sum(l.ILOC_BAL_QTY) as BAL_QTY " &
                        "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " &
                        "where ISNULL(ILOC_BAL_QTY, 0) > 0 " &
                        "and l.IMP_CODE = a.IMP_CODE " &
                        "and l.ILOC_WH = a.WH_CODE " &
                        "and l.ILOC_FLOOR = a.FL_NUM " &
                        "and l.ILOC_AREA = a.AR_CODE " &
                        "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " &
                        "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000'), l.ILOC_BATCH_NO) bal, " &
                        "( " &
                        "select l.imp_code, " &
                            "l.storer_code, " &
                            "l.ITM_CODE, " &
                            "l.PACK_KEY, " &
                            "ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                            "sum(l.ILOC_BAL_QTY) as BAL_QTY " &
                        "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " &
                        "where ISNULL(l.ILOC_BAL_QTY, 0) > 0 " &
                        "and l.IMP_CODE = a.IMP_CODE " &
                        "and l.ILOC_WH = a.WH_CODE " &
                        "and l.ILOC_FLOOR = a.FL_NUM " &
                        "and l.ILOC_AREA = a.AR_CODE " &
                        "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " &
                        "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000')) bal2 " &
                    "where i.imp_code = delv.imp_code " &
                    "and i.storer_code = delv.storer_code " &
                    "and i.itm_code = delv.pld_item_no " &
                    "and i.pack_key = delv.pld_pack_key " &
                    "and delv.imp_code = hold.imp_code(+) " &
                    "and delv.storer_code = hold.storer_code(+) " &
                    "and delv.pld_item_no = hold.cod_itm_code(+) " &
                    "and delv.pld_pack_key = hold.cod_pack_key(+) " &
                    "and delv.pld_pallet_no = hold.cod_pallet_no(+) " &
                    "and delv.pld_batch_no = hold.cod_batch_no(+) " &
                    "and delv.imp_code = hold2.imp_code(+) " &
                    "and delv.storer_code = hold2.storer_code(+) " &
                    "and delv.pld_item_no = hold2.cod_itm_code(+) " &
                    "and delv.pld_pack_key = hold2.cod_pack_key(+) " &
                    "and delv.pld_pallet_no = hold2.cod_pallet_no(+) " &
                    "and delv.imp_code = bal.imp_code(+) " &
                    "and delv.storer_code = bal.storer_code(+) " &
                    "and delv.pld_item_no = bal.ITM_CODE(+) " &
                    "and delv.pld_pack_key = bal.PACK_KEY(+) " &
                    "and delv.pld_pallet_no = bal.iloc_pallet_no(+) " &
                    "and delv.pld_batch_no = bal.iloc_batch_no(+) " &
                    "and delv.imp_code = bal2.imp_code(+) " &
                    "and delv.storer_code = bal2.storer_code(+) " &
                    "and delv.pld_item_no = bal2.ITM_CODE(+) " &
                    "and delv.pld_pack_key = bal2.PACK_KEY(+) " &
                    "and delv.pld_pallet_no = bal2.iloc_pallet_no(+) " &
                    "and ISNULL(Case delv.pld_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) - " &
                        "ISNULL(case delv.pld_batch_no when null then hold2.hold_qty else hold.hold_qty end, " &
                        "0) < delv.pld_item_qty " &
                    "order by delv.pld_item_no "

        '"group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold, " & _

        tmpDt = gDB.getDataTable(selectSql, , , , cmdPa)

        If tmpDt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                alertMsg = "Picked qty is greater than available qty, please check the stock balance of below items:"
            Else
                alertMsg = "實取貨量大於可取貨量, 請檢查以下貨物的可取存量:"
            End If

            For i = 0 To tmpDt.Rows.Count - 1
                alertMsg = alertMsg & vbNewLine &
                           "Item: " & tmpDt.Rows(i).Item("itm_code").ToString.Trim & ", Picked: " & tmpDt.Rows(i).Item("pld_item_qty").ToString.Trim & ", Available: " & tmpDt.Rows(i).Item("avail_qty").ToString.Trim
            Next

            'alertMsg = "test"

            uiFun.displayMsgNew(Me, "", alertMsg, Session("gLang"))

            Return False
        End If


        cmdPa = New GlobalDBFunc.DBCmdPara

        selectSql = "select d.pld_item_no, d.pld_loc, ISNULL(d.pld_item_qty, 0) as pick_qty, ISNULL(l.iloc_bal_qty, 0) as bal_qty " &
                    "from wms_do_picklist_d d, wms_item_loc_bal l " &
                    "where d.imp_code = l.imp_code " &
                    "and d.storer_code = l.storer_code " &
                    "and d.pld_item_no = l.itm_code " &
                    "and d.pld_pack_key = l.pack_key " &
                    "and ISNULL(d.pld_pallet_no, '000') = ISNULL(l.iloc_pallet_no, '000') " &
                    "and ISNULL(d.pld_batch_no, '') = ISNULL(l.iloc_batch_no, '') " &
                    "and d.pld_loc = l.iloc_loc " &
                    "and ISNULL(d.pld_item_qty, 0) > ISNULL(l.iloc_bal_qty, 0) " &
                    "and d.imp_code = " & cmdPa.AP(Session("IMP_CODE")) & " " &
                    "and d.storer_code = " & cmdPa.AP(STORER_CODE.SelectedValue) & " " &
                    "and d.do_code = " & cmdPa.AP(DO_CODE.Text.Trim) & " " &
                    "order by d.pld_item_no, d.pld_loc "

        tmpDt = gDB.getDataTable(selectSql, , , , cmdPa)

        If tmpDt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                alertMsg = "Picked qty is greater than balance qty, please check the stock balance of below items:"
            Else
                alertMsg = "實取貨量大於貨存量, 請檢查以下貨物的存量:"
            End If

            For i = 0 To tmpDt.Rows.Count - 1
                alertMsg = alertMsg & vbNewLine &
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
                    sql_string = "insert into WMS_DELV_ORDER " &
                    "(IMP_CODE, STORER_CODE, DO_CODE, DO_STATUS, DO_ISSUED_BY, DO_CO_CODE, DO_CUS_REF_NO, DO_PROJECT_NO, DO_DATE, DO_TARGET_DELDATE, " &
                    "DO_CONF_DELDATE, DO_CONF_DELTIME, CUS_CODE, CUS_NAME, DO_ADDR1, DO_ADDR2, DO_ADDR3, DO_AREA_DEL, DO_REGION_DEL, " &
                    "DO_COUNTRY_DEL, DO_CUS_CONT, DO_CUS_CONT_TEL, DO_DRIVER, DO_DRIVER_TEL, DO_VEHICLE_NO, DO_TOTL_PALLETS, " &
                    "DO_TOTL_CARTONS, DO_TOTL_BINS, DO_REM, DO_FTRACK_NO, " &
                    "DO_CONSIGNEE, DO_CONSIGNEE_ADDR1, DO_CONSIGNEE_ADDR2, DO_CONSIGNEE_ADDR3, DO_SHIP_TO, DO_SHIP_ADDR1,DO_SHIP_ADDR2,DO_SHIP_ADDR3, " &
                    "DO_PAY_TERMS, DO_TRADE_TERMS, DO_SHIP_MODE, DO_INV_NO, DO_PACK_LABEL_1, DO_PACK_LABEL_2, DO_PACK_LABEL_3, DO_PACK_LABEL_QTY_1, " &
                    "DO_PACK_LABEL_QTY_2, DO_PACK_LABEL_QTY_3, DO_TRANS_TYPE,DO_EDI_SIR_NO, DO_TROLLEY_ID, DO_DRUM_ID, DO_EDI_WIT_NO, " &
                    "DO_DELIVERY_RMKS, DO_ARRIVAL_DATE, DO_POST_CODE,DO_PROVINCE, DO_CITY,  " &
                    "DO_TO_STORER_REM, DO_STORER_REM,DO_SENDER,ROUTE_ID,DO_SENDER_COUNTRY, DO_SENDER_PROVINCE, DO_SENDER_REGION, DO_SENDER_ADDR, DO_SENDER_TEL," &
                     "sys_cb, sys_cd, sys_lub, sys_lud) " &
                    "values ( " &
                    "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', '" & gU.dbEncode(DO_STATUS.Text.Trim) & "', " &
                    "'" & gU.dbEncode(DO_ISSUED_BY.Text.Trim) & "', '" & gU.dbEncode(DO_CO_CODE.Text.Trim) & "', '" & gU.dbEncode(DO_CUS_REF_NO.Text.Trim) & "', " &
                    "'" & gU.dbEncode(DO_PROJECT_NO.SelectedValue) & "', " & gU.convdbDate(gU.dbEncode(DO_DATE.Text.Trim)) & ", " & gU.convdbDate(gU.dbEncode(DO_TARGET_DELDATE.Text.Trim)) & ", " &
                    gU.convdbDate(gU.dbEncode(DO_CONF_DELDATE.Text.Trim)) & ", "

                    If DO_CONF_DELDATE.Text.Trim <> "" AndAlso DO_CONF_DELTIME.Text.Trim <> "" Then
                        sql_string = sql_string &
                            "to_date('" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', '" & gU.getConfig("DDFORMAT") & " hh24:mi'), "
                    Else
                        sql_string = sql_string & "null, "
                    End If

                    sql_string = sql_string &
                    "'" & gU.dbEncode(CUS_CODE.SelectedValue) & "', " & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_ADDR1.Text.Trim)) & ", " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_ADDR2.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_ADDR3.Text.Trim)) & ", " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_AREA_DEL.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_REGION_DEL.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_COUNTRY_DEL.Text.Trim)) & ", " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_CUS_CONT.Text.Trim)) & ", '" & gU.dbEncode(DO_CUS_CONT_TEL.Text.Trim) & "', " & gU.convdbNVCData(gU.dbEncode(DO_DRIVER.Text.Trim)) & ", " &
                    "'" & gU.dbEncode(DO_DRIVER_TEL.Text.Trim) & "', '" & gU.dbEncode(DO_VEHICLE_NO.Text.Trim) & "', " &
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_PALLETS.Text.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_CARTONS.Text.Trim, "0")) & "', " &
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_BINS.Text.Trim, "0")) & "', " & gU.convdbNVCData(gU.dbEncode(DO_REM.Text.Trim)) & ",'" & gU.dbEncode(DO_FTRACK_NO.Text.Trim) & "', " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR1.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR2.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR3.Text.Trim)) & ", " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_SHIP_TO.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR1.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR2.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR3.Text.Trim)) & ", " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_PAY_TERMS.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(DO_TRADE_TERMS.Text.Trim)) & ", " & "'" & gU.dbEncode(DO_SHIP_MODE.SelectedValue) & "', " & "'" & gU.dbEncode(DO_INV_NO.Text.Trim) & "', " &
                    "'" & gU.dbEncode(DO_PACK_LABEL_1.Value.Trim) & "', " & "'" & gU.dbEncode(DO_PACK_LABEL_2.Value.Trim) & "', " & "'" & gU.dbEncode(DO_PACK_LABEL_3.Value.Trim) & "', " &
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_1.Value.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_2.Value.Trim, "0")) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_3.Value.Trim, "0")) & "', " &
                    "'" & gU.dbEncode(DO_TRANS_TYPE.SelectedValue) & "', " & "'" & gU.dbEncode(DO_EDI_SIR_NO.Text.Trim) & "', " &
                    "'" & gU.dbEncode(DO_TROLLEY_ID.Text.Trim) & "', " & "'" & gU.dbEncode(DO_DRUM_ID.Text.Trim) & "', " & "'" & gU.dbEncode(DO_EDI_WIT_NO.Text.Trim) & "', " &
                    "" & gU.convdbNVCData(gU.dbEncode(DO_DELIVERY_RMKS.Text.Trim)) & ", " & gU.convdbDate(gU.dbEncode(DO_ARRIVAL_DATE.Text.Trim)) & ", " & "'" & gU.dbEncode(DO_POST_CODE.Text.Trim) & "', " &
                    gU.convdbNVCData(gU.dbEncode(DO_PROVINCE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_CITY.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(DO_TO_STORER_REM.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_STORER_REM.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(DO_SENDER.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(ROUTE_ID.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(DO_SENDER_COUNTRY.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_SENDER_PROVINCE.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(DO_SENDER_REGION.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_SENDER_ADDR.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(DO_SENDER_TEL.Text.Trim)) & "," &
                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '"'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & nextNo & "', " & _
                    'DO_track_no,
                    ' "DO_TRACK_NO = '" & gU.dbEncode(DO_TRACK_NO.Text.Trim) & "', " & _
                    '" & gU.dbEncode(DO_TRACK_NO.Text.Trim) & "', 
                    REM **********************

                    REM Generate Document Link 
                    'dl.genDocLink("DO", DO_CODE.Text.Trim, "CO", DO_CO_CODE.text, STORER_CODE.SelectedValue, gConn, transaction)
                    dl.genDocLink("DO", nextNo, "CO", DO_CO_CODE.Text, STORER_CODE.SelectedValue, gConn, transaction)

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                        'uiFun.reOrderDetails(dt, "dod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into WMS_DELV_ORDER_D " &
                                            "(IMP_CODE, STORER_CODE, DO_CODE, DOD_SEQ, DOD_DISP_SEQ, DOD_PALLET_NO, DOD_CARTON_NO, DOD_PACK_NO, DOD_ITM_CODE, DOD_PACK_KEY,DOD_RETURN_QTY, " &
                                            "DOD_ITM_DESC, DOD_PACK_TYPE, DOD_QTY, DOD_UOM, DOD_CUT_YN, DOD_QTY2, DOD_UOM2, DOD_PCS_UOM, DOD_TOTPCS, DOD_TOT_WGT, DOD_TOT_CBM," &
                                            "DOD_REM, DOD_TICKET_NO, DOD_WH_CODE,DOD_FL_CODE,DOD_LOC_WH, " &
                                            "DOD_VND_CODE,DOD_BATCH_NO, DOD_EXPIRY_DATE, DOD_MANU_DATE, " &
                                            "DOD_PK_QTY,DOD_SS_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT," &
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                            "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_return_qty").ToString.Trim, "0")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_cut_yn").ToString.Trim, "N")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_WH_CODE").ToString.Trim, "")) & "', " &
                                               "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_FL_CODE").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LOC_WH").ToString.Trim, "")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                             "N'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_PK_QTY").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_SS_QTY").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_PROD_DATE").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_SHELF_LIFE").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LAST_LOT").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_MAX_LOT").ToString.Trim, "")) & "', " &
                                             "N'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"

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
                                itemSQL = "insert into wms_do_packing_d " &
                                        "(imp_code,storer_code,do_code," &
                                        "pad_pack_no,pad_pallet_no,pad_pack_type," &
                                        "pad_totl_packs, pad_uom, pad_carton_no, " &
                                        "pad_pack_key, pad_itm_code, pad_vnd_code, " &
                                        "pad_batch_no, pad_qty, pad_length, " &
                                        "pad_width, pad_height, pad_ref_no, " &
                                        "pad_pack_by, pad_display_seq, pad_net_weight, pad_gross_weight, pad_min_packing, pad_origin,pad_pack_size, PAD_QTY_PER_CTN, " &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " &
                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, ""))) & ", " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, "0")) & "', " &
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
                                itemSQL = "insert into WMS_DO_PICKLIST_D " &
                                        "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " &
                                        "PLD_ITEM_QTY,PLD_SS_QTY, PLD_WH, PLD_LOC, PLD_ORG_LOC, PLD_REMARK, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " &
                                        "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, PLD_TO_DRUM_ID, PLD_TO_DRUM_CABLE_LIST,PLD_TO_DRUM_ILOC_SEQ, " &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                          "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_ss_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_org_loc").ToString.Trim, "")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_remark").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " &
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, ""))) & ", " &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                '"('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.Text) & "', '" & nextNo & "', " & _
                        End Select
                        REM **********************
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    If pl_dt IsNot Nothing AndAlso pl_dt.Rows.Count > 0 Then
                        selectSql = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(nextNo) & "' "

                        plSeq = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSql, gConn, transaction), 0)

                        'For new added item, generate a empty picklist automatically
                        selectSql = "select DOD_ITM_CODE, DOD_PACK_KEY, DOD_PALLET_NO, DOD_QTY, DOD_BATCH_NO, DOD_QTY2, " &
                                        "Convert(varchar,DOD_EXPIRY_DATE, " & gU.getConfig("DDFORMATNO") & ") as DOD_EXPIRY_DATE, " &
                                        "Convert(varchar,DOD_MANU_DATE, " & gU.getConfig("DDFORMATNO") & ") as DOD_MANU_DATE," &
                                    "from WMS_DELV_ORDER_D d " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and DO_CODE = '" & gU.dbEncode(nextNo) & "' " &
                                    "and not exists ( " &
                                        "select 1 from WMS_DO_PICKLIST_D p " &
                                        "where p.IMP_CODE = d.IMP_CODE " &
                                        "and p.STORER_CODE = d.STORER_CODE " &
                                        "and p.DO_CODE = d.DO_CODE " &
                                        "and p.PLD_ITEM_NO = d.DOD_ITM_CODE " &
                                        "and p.PLD_PACK_KEY = d.DOD_PACK_KEY) "

                        mplDt = gDB.getDataTable(selectSql, gConn, transaction)

                        'plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                        For i = 0 To mplDt.Rows.Count - 1
                            plSeq = plSeq + 1

                            If gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, "")) <> "" Then
                                itemSQL = "insert into WMS_DO_PICKLIST_D " &
                                            "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_ITEM_QTY, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " &
                                             "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, " &
                                             "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                             "'" & gU.dbEncode(CStr(plSeq)) & "', '" & Session("usr_id") & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim, "000")) & "', " &
                                             "0, '', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("dod_qty").ToString.Trim, "0")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(mplDt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                             "0, " &
                                             gU.convdbDate(gU.dbEncode(mplDt.Rows(i).Item("DOD_EXPIRY_DATE").ToString.Trim)) & ", " &
                                             gU.convdbDate(gU.dbEncode(mplDt.Rows(i).Item("DOD_MANU_DATE").ToString.Trim)) & ", " &
                                             "null, " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(mplDt.Rows(i).Item("dod_qty2").ToString.Trim, "0")) & "', " &
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                gDB.amendData(itemSQL, gConn, transaction)
                            End If
                        Next

                        Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)
                    End If


                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode

                    sql_string = "update WMS_DELV_ORDER set " &
                                    "DO_ISSUED_BY = '" & gU.dbEncode(DO_ISSUED_BY.Text.Trim) & "', " &
                                    "DO_CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text.Trim) & "', " &
                                    "DO_CUS_REF_NO = '" & gU.dbEncode(DO_CUS_REF_NO.Text.Trim) & "', " &
                                    "DO_PROJECT_NO = '" & gU.dbEncode(DO_PROJECT_NO.SelectedValue) & "', " &
                                    "DO_DATE = " & gU.convdbDate(gU.dbEncode(DO_DATE.Text.Trim)) & ", " &
                                    "DO_TARGET_DELDATE = " & gU.convdbDate(gU.dbEncode(DO_TARGET_DELDATE.Text.Trim)) & ", " &
                                    "DO_CONF_DELDATE = " & gU.convdbDate(gU.dbEncode(DO_CONF_DELDATE.Text.Trim)) & ", " &
                                    "DO_CONF_DELTIME = "

                    If DO_CONF_DELDATE.Text.Trim <> "" AndAlso DO_CONF_DELTIME.Text.Trim <> "" Then
                        sql_string = sql_string &
                            "to_date('" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', '" & gU.getConfig("DDFORMAT") & " hh24:mi'), "
                    Else
                        sql_string = sql_string & "null, "
                    End If

                    sql_string = sql_string &
                                    "CUS_CODE = '" & gU.dbEncode(CUS_CODE.SelectedValue) & "', " &
                                    "CUS_NAME = " & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text.Trim)) & ", " &
                                    "DO_ADDR1 = " & gU.convdbNVCData(gU.dbEncode(DO_ADDR1.Text.Trim)) & ", " &
                                    "DO_ADDR2 = " & gU.convdbNVCData(gU.dbEncode(DO_ADDR2.Text.Trim)) & ", " &
                                    "DO_ADDR3 = " & gU.convdbNVCData(gU.dbEncode(DO_ADDR3.Text.Trim)) & ", " &
                                    "DO_AREA_DEL = " & gU.convdbNVCData(gU.dbEncode(DO_AREA_DEL.Text.Trim)) & ", " &
                                    "DO_REGION_DEL = " & gU.convdbNVCData(gU.dbEncode(DO_REGION_DEL.Text.Trim)) & ", " &
                                    "DO_COUNTRY_DEL = " & gU.convdbNVCData(gU.dbEncode(DO_COUNTRY_DEL.Text.Trim)) & ", " &
                                    "DO_CUS_CONT = " & gU.convdbNVCData(gU.dbEncode(DO_CUS_CONT.Text.Trim)) & ", " &
                                    "DO_CUS_CONT_TEL = '" & gU.dbEncode(DO_CUS_CONT_TEL.Text.Trim) & "', " &
                                    "DO_DRIVER = " & gU.convdbNVCData(gU.dbEncode(DO_DRIVER.Text.Trim)) & ", " &
                                    "DO_DRIVER_TEL = '" & gU.dbEncode(DO_DRIVER_TEL.Text.Trim) & "', " &
                                    "DO_VEHICLE_NO = '" & gU.dbEncode(DO_VEHICLE_NO.Text.Trim) & "', " &
                                    "DO_TOTL_PALLETS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_PALLETS.Text.Trim, "0")) & "', " &
                                    "DO_TOTL_CARTONS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_CARTONS.Text.Trim, "0")) & "', " &
                                    "DO_TOTL_BINS = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_TOTL_BINS.Text.Trim, "0")) & "', " &
                                    "DO_REM = " & gU.convdbNVCData(gU.dbEncode(DO_REM.Text.Trim)) & ", " &
                                    "DO_TO_STORER_REM = " & gU.convdbNVCData(gU.dbEncode(DO_TO_STORER_REM.Text.Trim)) & ", " &
                                    "DO_STORER_REM = " & gU.convdbNVCData(gU.dbEncode(DO_STORER_REM.Text.Trim)) & ", " &
                                    "DO_FTRACK_NO = '" & gU.dbEncode(DO_FTRACK_NO.Text.Trim) & "', " &
                                    "DO_CONSIGNEE = " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE.Text.Trim)) & ", " &
                                    "DO_CONSIGNEE_ADDR1 = " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR1.Text.Trim)) & ", " &
                                    "DO_CONSIGNEE_ADDR2 = " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR2.Text.Trim)) & ", " &
                                    "DO_CONSIGNEE_ADDR3 = " & gU.convdbNVCData(gU.dbEncode(DO_CONSIGNEE_ADDR3.Text.Trim)) & ", " &
                                    "DO_SHIP_TO = " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_TO.Text.Trim)) & ", " &
                                    "DO_SHIP_ADDR1 = " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR1.Text.Trim)) & ", " &
                                    "DO_SHIP_ADDR2 = " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR2.Text.Trim)) & ", " &
                                    "DO_SHIP_ADDR3 = " & gU.convdbNVCData(gU.dbEncode(DO_SHIP_ADDR3.Text.Trim)) & ", " &
                                    "DO_PAY_TERMS = " & gU.convdbNVCData(gU.dbEncode(DO_PAY_TERMS.Text.Trim)) & ", " &
                                    "DO_TRADE_TERMS = " & gU.convdbNVCData(gU.dbEncode(DO_TRADE_TERMS.Text.Trim)) & ", " &
                                    "DO_SHIP_MODE = '" & gU.dbEncode(DO_SHIP_MODE.SelectedValue) & "', " &
                                    "DO_INV_NO = '" & gU.dbEncode(DO_INV_NO.Text.Trim) & "', " &
                                    "DO_PACK_LABEL_1 = '" & gU.dbEncode(DO_PACK_LABEL_1.Value.Trim) & "', " &
                                    "DO_PACK_LABEL_2 = '" & gU.dbEncode(DO_PACK_LABEL_2.Value.Trim) & "', " &
                                    "DO_PACK_LABEL_3 = '" & gU.dbEncode(DO_PACK_LABEL_3.Value.Trim) & "', " &
                                    "DO_PACK_LABEL_QTY_1 = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_1.Value.Trim, "0")) & "', " &
                                    "DO_PACK_LABEL_QTY_2 = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_2.Value.Trim, "0")) & "', " &
                                    "DO_PACK_LABEL_QTY_3 = '" & gU.dbEncode(gU.decodeNullOrEmpty(DO_PACK_LABEL_QTY_3.Value.Trim, "0")) & "', " &
                                    "DO_TRANS_TYPE = " & "'" & gU.dbEncode(DO_TRANS_TYPE.SelectedValue) & "', " &
                                    "DO_EDI_SIR_NO = " & "'" & gU.dbEncode(DO_EDI_SIR_NO.Text.Trim) & "', " &
                                    "DO_TROLLEY_ID = " & "'" & gU.dbEncode(DO_TROLLEY_ID.Text.Trim) & "', " &
                                    "DO_DRUM_ID = " & "'" & gU.dbEncode(DO_DRUM_ID.Text.Trim) & "', " &
                                    "DO_EDI_WIT_NO = " & "'" & gU.dbEncode(DO_EDI_WIT_NO.Text.Trim) & "', " &
                                    "DO_DELIVERY_RMKS = " & gU.convdbNVCData(gU.dbEncode(DO_DELIVERY_RMKS.Text.Trim)) & ", " &
                                    "DO_ARRIVAL_DATE = " & gU.convdbDate(gU.dbEncode(DO_ARRIVAL_DATE.Text.Trim)) & ", " &
                                    "DO_POST_CODE = " & "'" & gU.dbEncode(DO_POST_CODE.Text.Trim) & "', " &
                                    "DO_PROVINCE=" & gU.convdbNVCData(gU.dbEncode(DO_PROVINCE.Text.Trim)) & ", " &
                                    "DO_CITY=" & gU.convdbNVCData(gU.dbEncode(DO_CITY.Text.Trim)) & ", " &
                                    "DO_SENDER=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER.Text.Trim)) & ", " &
                                    "ROUTE_ID=" & gU.convdbNVCData(gU.dbEncode(ROUTE_ID.Text.Trim)) & ", " &
                                    "DO_SENDER_COUNTRY=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER_COUNTRY.Text.Trim)) & ", " &
                                    "DO_SENDER_PROVINCE=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER_PROVINCE.Text.Trim)) & ", " &
                                    "DO_SENDER_REGION=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER_REGION.Text.Trim)) & ", " &
                                    "DO_SENDER_ADDR=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER_ADDR.Text.Trim)) & ", " &
                                    "DO_SENDER_TEL=" & gU.convdbNVCData(gU.dbEncode(DO_SENDER_TEL.Text.Trim)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                        'uiFun.reOrderDetails(dt, "dod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag").ToString
                                Case "N"
                                    itemSQL = "insert into WMS_DELV_ORDER_D " &
                                            "(IMP_CODE, STORER_CODE, DO_CODE, DOD_SEQ, DOD_DISP_SEQ, DOD_PALLET_NO, DOD_CARTON_NO, DOD_PACK_NO, DOD_ITM_CODE, DOD_PACK_KEY,DOD_RETURN_QTY, " &
                                            "DOD_ITM_DESC, DOD_PACK_TYPE, DOD_QTY, DOD_UOM, DOD_CUT_YN, DOD_QTY2, DOD_UOM2, DOD_PCS_UOM, DOD_TOTPCS, DOD_TOT_WGT, DOD_TOT_CBM, " &
                                            "DOD_REM, DOD_TICKET_NO, DOD_WH_CODE,DOD_FL_CODE,DOD_LOC_WH, " &
                                            "DOD_VND_CODE,DOD_BATCH_NO, DOD_EXPIRY_DATE, DOD_MANU_DATE, " &
                                            "DOD_PK_QTY,DOD_SS_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT," &
                                             "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                            "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_return_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_cut_yn").ToString.Trim, "N")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_WH_CODE").ToString.Trim, "")) & "', " &
                                              "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_FL_CODE").ToString.Trim, "")) & "', " &
                                              "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LOC_WH").ToString.Trim, "")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                             "N'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_PK_QTY").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_SS_QTY").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_PROD_DATE").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_SHELF_LIFE").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LAST_LOT").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_MAX_LOT").ToString.Trim, "")) & "', " &
                                             "N'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from WMS_DELV_ORDER_D " &
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                            "and DOD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update WMS_DELV_ORDER_D set " &
                                                "DOD_DISP_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_disp_seq").ToString.Trim, "")) & "', " &
                                                "DOD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "000")) & "', " &
                                                "DOD_CARTON_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_carton_no").ToString.Trim, "")) & "', " &
                                                "DOD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_no").ToString.Trim, "")) & "', " &
                                                "DOD_RETURN_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_return_qty").ToString.Trim, "0")) & "', " &
                                                "DOD_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "")) & "', " &
                                                "DOD_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "")) & "', " &
                                                "DOD_ITM_DESC = N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_itm_desc").ToString.Trim, "")) & "', " &
                                                "DOD_PACK_TYPE = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_pack_type").ToString.Trim, "")) & "', " &
                                                "DOD_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")) & "', " &
                                                "DOD_UOM = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom").ToString.Trim, "")) & "', " &
                                                "DOD_CUT_YN = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_CUT_YN").ToString.Trim, "N")) & "', " &
                                                "DOD_QTY2 = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) & "', " &
                                                "DOD_UOM2 = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_uom2").ToString.Trim, "")) & "', " &
                                                "DOD_PCS_UOM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_pcs_uom").ToString.Trim, "0")) & "', " &
                                                "DOD_TOTPCS = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_totpcs").ToString.Trim, "0")) & "', " &
                                                "DOD_TOT_WGT = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_wgt").ToString.Trim, "0")) & "', " &
                                                "DOD_TOT_CBM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("dod_tot_cbm").ToString.Trim, "0")) & "', " &
                                                "DOD_REM = N'" & gU.dbEncode(gU.decodeNull(rows.Item("dod_rem").ToString.Trim, "")) & "', " &
                                                "DOD_TICKET_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_TICKET_NO").ToString.Trim, "")) & "', " &
                                                "DOD_WH_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_WH_CODE").ToString.Trim, "")) & "', " &
                                                "DOD_FL_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_fl_code").ToString.Trim, "")) & "', " &
                                                "DOD_LOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LOC_WH").ToString.Trim, "")) & "', " &
                                                "dod_vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("dod_vnd_code").ToString.Trim, "")) & "', " &
                                                "DOD_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("DOD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                                "DOD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                "DOD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("DOD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                "DOD_PK_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_PK_QTY").ToString.Trim, "0")) & "', " &
                                                "DOD_SS_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_SS_QTY").ToString.Trim, "0")) & "', " &
                                                "DOD_MIN_PROD_DATE = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_PROD_DATE").ToString.Trim, "0")) & "', " &
                                                "DOD_MIN_SHELF_LIFE ='" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DOD_MIN_SHELF_LIFE").ToString.Trim, "0")) & "', " &
                                                "DOD_LAST_LOT = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_LAST_LOT").ToString.Trim, "")) & "', " &
                                                "DOD_MAX_LOT = '" & gU.dbEncode(gU.decodeNull(rows.Item("DOD_MAX_LOT").ToString.Trim, "")) & "', " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
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
                                itemSQL = "insert into wms_do_packing_d " &
                                       "(imp_code,storer_code,do_code," &
                                       "pad_pack_no,pad_pallet_no,pad_pack_type," &
                                       "pad_totl_packs, pad_uom, pad_carton_no, " &
                                       "pad_pack_key, pad_itm_code, pad_vnd_code, " &
                                       "pad_batch_no, pad_qty, pad_length, " &
                                       "pad_width, pad_height, pad_ref_no, " &
                                       "pad_pack_by, pad_display_seq, pad_net_weight, pad_gross_weight, pad_min_packing, pad_origin,pad_pack_size,PAD_QTY_PER_CTN," &
                                       "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                       "values " &
                                       "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " &
                                       gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, ""))) & ", " &
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " &
                                       "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " &
                                        "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, "0")) & "', " &
                                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "


                            Case "D", "R"
                                itemSQL = "delete from WMS_DO_PACKING_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                        "and PAD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_no").ToString.Trim, "")) & "' "
                            Case Else
                                itemSQL = "update WMS_DO_PACKING_D set " &
                                            "PAD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pallet_no").ToString.Trim, "000")) & "', " &
                                            "PAD_PACK_TYPE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_type").ToString.Trim, "")) & "', " &
                                            "PAD_TOTL_PACKS = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_totl_packs").ToString.Trim, "0")) & "', " &
                                            "pad_uom = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_uom").ToString.Trim, "")) & "', " &
                                           "pad_carton_no = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_carton_no").ToString.Trim, "")) & "', " &
                                           "pad_pack_key = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_key").ToString.Trim, "")) & "', " &
                                           "pad_pack_size = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_pack_size").ToString.Trim, "0")) & "', " &
                                           "pad_itm_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_itm_code").ToString.Trim, "")) & "', " &
                                           "pad_vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_vnd_code").ToString.Trim, "")) & "', " &
                                           "pad_batch_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pad_batch_no").ToString.Trim, ""))) & ", " &
                                           "pad_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_qty").ToString.Trim, "0")) & ", " &
                                           "pad_length = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_length").ToString.Trim, "0")) & ", " &
                                           "pad_width = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_width").ToString.Trim, "0")) & ", " &
                                           "pad_height = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_height").ToString.Trim, "0")) & ", " &
                                           "pad_ref_no = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_ref_no").ToString.Trim, "")) & "', " &
                                           "pad_pack_by = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_pack_by").ToString.Trim, "")) & "', " &
                                           "pad_display_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_display_seq").ToString.Trim, "")) & "', " &
                                           "pad_net_weight = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_net_weight").ToString.Trim, "0")) & "', " &
                                           "pad_gross_weight = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_gross_weight").ToString.Trim, "0")) & "', " &
                                           "pad_min_packing = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pad_min_packing").ToString.Trim, "0")) & "', " &
                                           "pad_origin = '" & gU.dbEncode(gU.decodeNull(rows.Item("pad_origin").ToString.Trim, "")) & "', " &
                                           "PAD_QTY_PER_CTN = '" & gU.dbEncode(gU.decodeEmptyCInt(rows.Item("PAD_QTY_PER_CTN").ToString.Trim, 0)) & "', " &
                                           "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
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
                            itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
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
                                itemSQL = "insert into WMS_DO_PICKLIST_D " &
                                        "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " &
                                        "PLD_ITEM_QTY,PLD_SS_QTY, PLD_WH, PLD_LOC, PLD_ORG_LOC, PLD_REMARK, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, " &
                                        "PLD_EXPIRY_DATE, PLD_MANU_DATE, PLD_SERIAL_NO, PLD_QTY2, PLD_TO_DRUM_ID, PLD_TO_DRUM_CABLE_LIST,PLD_TO_DRUM_ILOC_SEQ, " &
                                         "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(DO_CODE.Text) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_ss_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_org_loc").ToString.Trim, "")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_remark").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " &
                                         gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, ""))) & ", " &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            Case "D"
                                itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                        "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "' "
                            Case Else
                                itemSQL = "update WMS_DO_PICKLIST_D set " &
                                            "PLD_PICKED_BY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                            "PLD_ITEM_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "', " &
                                            "PLD_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "', " &
                                            "PLD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "', " &
                                            "PLD_ITEM_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                            "PLD_SS_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_ss_qty").ToString.Trim, "0")) & "', " &
                                            "PLD_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "', " &
                                            "PLD_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "', " &
                                            "PLD_ORG_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_org_loc").ToString.Trim, "")) & "', " &
                                            "PLD_REMARK = N'" & gU.dbEncode(gU.decodeNull(rows.Item("pld_remark").ToString.Trim, "")) & "', " &
                                            "PLD_FLOOR = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_floor").ToString.Trim, "")) & "', " &
                                            "PLD_AREA = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")) & "', " &
                                            "PLD_RACK = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_rack").ToString.Trim, "")) & "', " &
                                            "PLD_BIN = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_bin").ToString.Trim, "")) & "', " &
                                            "PLD_IS_LOAN = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                            "PLD_DO_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                            "PLD_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_BATCH_NO").ToString.Trim, ""))) & ", " &
                                            "PLD_FOI_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                            "PLD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(rows.Item("pld_expiry_date").ToString.Trim)) & ", " &
                                            "PLD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(rows.Item("pld_manu_date").ToString.Trim)) & ", " &
                                            "PLD_SERIAL_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_serial_no").ToString.Trim, "")) & "', " &
                                            "PLD_QTY2 = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_qty2").ToString.Trim, "0")) & "', " &
                                            "PLD_TO_DRUM_ID = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, ""))) & ", " &
                                            "PLD_TO_DRUM_CABLE_LIST = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, ""))) & ", " &
                                            "PLD_TO_DRUM_ILOC_SEQ = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, ""))) & ", " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                        "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "'"

                        End Select
                        REM **********************

                        'Response.Write(itemSQL)
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    'Delete item not selected anymore
                    itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                "and not exists ( " &
                                    "select 1 from WMS_DELV_ORDER_D d " &
                                    "where WMS_DO_PICKLIST_D.IMP_CODE = d.IMP_CODE " &
                                    "and WMS_DO_PICKLIST_D.STORER_CODE = d.STORER_CODE " &
                                    "and WMS_DO_PICKLIST_D.DO_CODE = d.DO_CODE " &
                                    "and WMS_DO_PICKLIST_D.PLD_ITEM_NO = d.DOD_ITM_CODE " &
                                    "and WMS_DO_PICKLIST_D.PLD_PACK_KEY = d.DOD_PACK_KEY) "

                    gDB.amendData(itemSQL, gConn, transaction)

                    'Delete item not selected anymore
                    itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                "and ISNULL(PLD_ITEM_QTY, 0) = 0 " &
                                "and not exists ( " &
                                    "select 1 from WMS_DELV_ORDER_D d " &
                                    "where WMS_DO_PICKLIST_D.IMP_CODE = d.IMP_CODE " &
                                    "and WMS_DO_PICKLIST_D.STORER_CODE = d.STORER_CODE " &
                                    "and WMS_DO_PICKLIST_D.DO_CODE = d.DO_CODE " &
                                    "and WMS_DO_PICKLIST_D.PLD_ITEM_NO = d.DOD_ITM_CODE " &
                                    "and WMS_DO_PICKLIST_D.PLD_PACK_KEY = d.DOD_PACK_KEY " &
                                    "and ISNULL(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') = ISNULL(d.DOD_PALLET_NO, '000')) "

                    gDB.amendData(itemSQL, gConn, transaction)

                    selectSql = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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
                    updateSql = "update WMS_DELV_ORDER " &
                                "set DO_STATUS = 'ASSIGNED' " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSql, gConn, transaction)

                    updateSql = "update WMS_DO_PICKLIST_D " &
                                "set PLD_STATUS = 'ASSIGNED' " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSql, gConn, transaction)
                Else

                    If DO_STATUS.Text = "ASSIGNED" Then
                        updateSql = "update WMS_DO_PICKLIST_D " &
                                    "set PLD_STATUS = 'ASSIGNED' " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                                    "and isnull(PLD_STATUS, 'NEW') = 'NEW' " &
                                    "and exists (" &
                                        "select 1 from WMS_DELV_ORDER d " &
                                        "where d.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " &
                                        "and d.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " &
                                        "and d.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " &
                                        "and d.DO_STATUS = 'ASSIGNED') "

                        gDB.amendData(updateSql, gConn, transaction)
                    End If

                End If

                updatePLParentChild(gConn, transaction)

                For Each rows As DataRow In dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                dt.AcceptChanges()

                For Each rows As DataRow In pi_dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                pi_dt.AcceptChanges()

                For Each rows As DataRow In pl_dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                pl_dt.AcceptChanges()

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

        updateSql = "update WMS_DO_PICKLIST_D " &
                    "set PLD_SPLIT_FR = PLD_SEQ " &
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

        '"and PLD_SPLIT_FR is not null "

        gDB.amendData(updateSql, gConn, transaction)

        updateSql = "update WMS_DO_PICKLIST_D " &
                    "set PLD_SPLIT_FR = (" &
                        "select convert(varchar, max(convert(int, p.PLD_SEQ))) " &
                        "from WMS_DO_PICKLIST_D p " &
                        "where p.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " &
                        "and p.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " &
                        "and p.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " &
                        "and p.PLD_ITEM_NO = WMS_DO_PICKLIST_D.PLD_ITEM_NO " &
                        "and p.PLD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " &
                        "and isnull(p.PLD_PALLET_NO, '000') = isnull(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') " &
                        "and isnull(p.PLD_BATCH_NO, '') = isnull(WMS_DO_PICKLIST_D.PLD_BATCH_NO, '') " &
                        "and convert(int, p.PLD_SEQ) < convert(int, WMS_DO_PICKLIST_D.PLD_SEQ)) " &
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' " &
                    "and exists (" &
                        "select 1 from WMS_DO_PICKLIST_D p " &
                        "where p.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " &
                        "and p.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE " &
                        "and p.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " &
                        "and p.PLD_ITEM_NO = WMS_DO_PICKLIST_D.PLD_ITEM_NO " &
                        "and p.PLD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " &
                        "and isnull(p.PLD_PALLET_NO, '000') = isnull(WMS_DO_PICKLIST_D.PLD_PALLET_NO, '000') " &
                        "and isnull(p.PLD_BATCH_NO, '') = isnull(WMS_DO_PICKLIST_D.PLD_BATCH_NO, '') " &
                        "and convert(int, p.PLD_SEQ) < convert(int, WMS_DO_PICKLIST_D.PLD_SEQ)) "

        gDB.amendData(updateSql, gConn, transaction)

    End Sub

    Protected Sub BindGV()

        Dim SQLString As String = ""
        Dim dt, dtV, dtT As New DataTable
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

            Session("DO_CODE") = doCode
            Session("STORER_CODE") = storerCode

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
            btnNOTE.Visible = False
            uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " &
                                          "storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                        "and imp_code = '" & Session("IMP_CODE") & "' " &
                                        "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

            uiFun.load_dropdown(DO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " &
                                          "storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                          "and imp_code = '" & Session("IMP_CODE") & "' " &
                                          "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
        Else
            REM **********************
            REM Modify Here
            btnAttach.Visible = True
            btnNOTE.Visible = True
            REM Generate Data Table from Header
            SQLString = "SELECT IMP_CODE, STORER_CODE, DO_CODE, DO_STATUS, DO_ISSUED_BY, DO_CO_CODE, DO_CUS_REF_NO, DO_PROJECT_NO, " &
                            "Convert(varchar, DO_DATE, " & gU.getConfig("DDFORMATNo") & ") as DO_DATE, " &
                            "Convert(varchar,DO_TARGET_DELDATE, " & gU.getConfig("DDFORMATNo") & ") as DO_TARGET_DELDATE, " &
                            "Convert(varchar,DO_CONF_DELDATE, " & gU.getConfig("DDFORMATNo") & ") as DO_CONF_DELDATE, " &
                            "Convert(varchar(5),DO_CONF_DELTIME, 108) as DO_CONF_DELTIME, " &
                            "CUS_CODE, CUS_NAME, DO_ADDR1, DO_ADDR2, DO_ADDR3, DO_AREA_DEL, " &
                            "DO_REGION_DEL, DO_COUNTRY_DEL, DO_CUS_CONT, DO_CUS_CONT_TEL, DO_DRIVER, DO_DRIVER_TEL, DO_VEHICLE_NO, DO_TOTL_PALLETS, " &
                            "DO_TOTL_CARTONS, DO_TOTL_BINS, DO_REM, DO_TO_STORER_REM, DO_STORER_REM, DO_FTRACK_NO, DO_POSTED_DATE, DO_POSTED_BY, DO_TEAM, DO_SCH_DATE, " &
                            "DO_CONTRACT_TYPE, DO_AGREEMENT_NO, DO_INSTALLATION_REQ, DO_INITIAL_COLLECT, DO_INITIAL_COLLECT_AMT, DO_PAY_TERMS, " &
                            "DO_DELIVERED_BY, DO_AREA_CODE, DO_DIS_CODE, DO_REGION_CODE, DO_DELI_STATUS, DO_TYPE, DO_ST_ID, DO_ST_SEQ, DO_BLOCK_DEL, " &
                            "DO_FLOOR_DEL, DO_FLAT_DEL, DO_RTE_CODE, DO_SIGNATURE, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, DO_JOB_ID, DO_PDASTATUS, " &
                            "DO_REASON_CODE, CUS_BILL_ADDR, CUST_CONT_NAME, CUST_CONT_PHONE, DO_ORG_ADDR, DO_ZONE, DO_SUB_ZONE, DO_ADDR_ID, " &
                            "DO_ADDR_SEQ, DO_SUB_TYPE, DO_ENT_CODE, DO_ENT_DESC, DO_DELI_INSTR, DO_TARGET_AMPM, DO_RESCH_DELDATE, DO_CMPLT_DELDATE, " &
                            "DO_COL_REMARK, DO_RCL_COD_AMT, DO_RCL_TONER_QTY, DO_RCL_CART_QTY, DO_RCL_PKG_QTY, DO_RCL_OTHER_QTY, DO_COL_COD_AMT, " &
                            "DO_COL_TONER_QTY, DO_COL_CART_QTY, DO_COL_PKG_QTY, DO_COL_OTHER_QTY, DO_ISSUED_DATE, DO_APPR_DATE, DO_APPR_BY, " &
                            "DO_CONF_DATE, DO_CONF_BY, DO_CANCEL_DATE, DO_CANCEL_BY, DO_DEL_DATE, DO_HOLD_DATE, DO_HOLD_BY, DO_RESCH_DATE, " &
                            "DO_RESCH_BY, DO_COMPL_DATE, DO_COMPL_BY, DO_VERIFY_DATE, DO_VERIFY_BY, DLAL_ID, DO_LEADER, DO_PDA, DO_SALES_NAME, " &
                            "DO_SALES_PHONE, DO_SALES_EMAIL, DO_MEM_TYPE, DO_DELI_TYPE, DO_PCK_ZONE, DO_PCK_SUB_ZONE, DO_PCK_ADDR_ID, " &
                            "DO_PCK_ADDR_SEQ, DO_PCK_FLAT, DO_PCK_FLOOR, DO_PCK_BLOCK, DO_SCAN_CODE, DO_INSTALL_INSTR, DO_INSTALL_DATE, " &
                            "DO_CODE_BF, DO_CODE_BF_TGT, DO_CODE_AF, DO_CODE_AF_TGT, DO_ORDER_NO, DO_RMA, DO_REF_DO_CODE, DO_SALES_CODE, DO_TARGET_DT_FR, " &
                            "DO_TARGET_DT_TO, DO_LATEST_DELI_DATE, DO_MEM_DOC_TYPE, DO_REF_NO, DO_CTRL_ENT, DO_WHS_CODE, DO_SELF_PICK_YN, " &
                            "DO_INSTALL_BY, DO_WT_TYPE, DO_MAP_ADDR, DO_REASON_DESC, DO_PDA_RS_CODE, DO_TERR_CODE, DO_ADDR_MAP_FLAG, DO_DAILY_SCH_FLAG, " &
                            "DO_DAILY_SCH_CHG_YN, DO_RCV_CUTOFF_YN, DO_SUBMIT_BY, DO_SUBMIT_DATE, DO_APPROVER, DO_REJ_BY, DO_REJ_ON, DO_REJ_REASON, " &
                            "DO_PCK_ADDR1, DO_MEMO_PRINT_YN, DO_MEMO_PRINT_DATE, DO_CONSIGNEE, DO_CONSIGNEE_ADDR1, DO_CONSIGNEE_ADDR2, " &
                            "DO_CONSIGNEE_ADDR3, DO_SHIP_TO, DO_SHIP_ADDR1, DO_SHIP_ADDR2, DO_SHIP_ADDR3, DO_INV_NO, DO_TRADE_TERMS, " &
                            "DO_SHIP_MODE, DO_PACK_LABEL_1, DO_PACK_LABEL_2, DO_PACK_LABEL_3, DO_PACK_LABEL_QTY_1, DO_PACK_LABEL_QTY_2, " &
                            "DO_PACK_LABEL_QTY_3, DO_TRANS_TYPE, DO_CONFIRM_DATE, DO_DELI_FLAG, DO_EDI_SIR_NO, DO_TROLLEY_ID, DO_DRUM_ID, DO_EDI_WIT_NO, DO_DELIVERY_RMKS, Convert(varchar,DO_ARRIVAL_DATE, " & gU.getConfig("DDFORMATNo") & ") as DO_ARRIVAL_DATE, DO_POST_CODE, " &
                            "DO_PROVINCE, DO_CITY,DO_SENDER,ROUTE_ID, DO_SENDER_COUNTRY, DO_SENDER_PROVINCE, DO_SENDER_REGION, DO_SENDER_ADDR, DO_SENDER_TEL " &
                        "from WMS_DELV_ORDER " &
                        "WHERE WMS_DELV_ORDER.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                        "AND WMS_DELV_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                        "AND WMS_DELV_ORDER.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            'DO_TRACK_NO,

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
                DO_TO_STORER_REM.Text = dt.Rows(0).Item("DO_TO_STORER_REM").ToString
                DO_STORER_REM.Text = dt.Rows(0).Item("DO_STORER_REM").ToString
                'DO_TRACK_NO.Text = dt.Rows(0).Item("DO_TRACK_NO").ToString
                'hide_DO_TRACK_NO.Value = dt.Rows(0).Item("DO_TRACK_NO").ToString
                DO_FTRACK_NO.Text = dt.Rows(0).Item("DO_FTRACK_NO").ToString
                'NEW FIELD

                DO_CONSIGNEE.Text = dt.Rows(0).Item("DO_CONSIGNEE").ToString
                DO_CONSIGNEE_ADDR1.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR1").ToString
                DO_CONSIGNEE_ADDR2.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR2").ToString
                DO_CONSIGNEE_ADDR3.Text = dt.Rows(0).Item("DO_CONSIGNEE_ADDR3").ToString
                DO_SHIP_TO.Text = dt.Rows(0).Item("DO_SHIP_TO").ToString
                DO_DELIVERY_RMKS.Text = dt.Rows(0).Item("DO_DELIVERY_RMKS").ToString
                DO_ARRIVAL_DATE.Text = dt.Rows(0).Item("DO_ARRIVAL_DATE").ToString
                DO_POST_CODE.Text = dt.Rows(0).Item("DO_POST_CODE").ToString
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
                DO_EDI_WIT_NO.Text = dt.Rows(0).Item("DO_EDI_WIT_NO").ToString.Trim
                DO_PROVINCE.Text = dt.Rows(0).Item("DO_PROVINCE").ToString.Trim
                DO_CITY.Text = dt.Rows(0).Item("DO_CITY").ToString.Trim
                DO_SENDER.Text = dt.Rows(0).Item("DO_SENDER").ToString.Trim
                ROUTE_ID.Text = dt.Rows(0).Item("ROUTE_ID").ToString.Trim
                DO_SENDER_COUNTRY.Text = dt.Rows(0).Item("DO_SENDER_COUNTRY").ToString.Trim
                DO_SENDER_PROVINCE.Text = dt.Rows(0).Item("DO_SENDER_PROVINCE").ToString.Trim
                DO_SENDER_REGION.Text = dt.Rows(0).Item("DO_SENDER_REGION").ToString.Trim
                DO_SENDER_ADDR.Text = dt.Rows(0).Item("DO_SENDER_ADDR").ToString.Trim
                DO_SENDER_TEL.Text = dt.Rows(0).Item("DO_SENDER_TEL").ToString.Trim

                'END HERE
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                Select Case dt.Rows(0).Item("DO_STATUS").ToString
                    Case "NEW"
                        'btnPost.Visible = False
                        btnPostChk.Visible = False
                        btnUnPost.Visible = False
                        btnPick.Visible = True
                        btnUnPick.Visible = False
                    Case "PICKED"
                        'btnPost.Visible = True
                        btnPostChk.Visible = True
                        btnUnPost.Visible = False
                        btnPick.Visible = False
                        btnUnPick.Visible = True
                    Case "POSTED"
                        'btnPost.Visible = False
                        btnPostChk.Visible = False
                        btnUnPost.Visible = True
                        btnPick.Visible = False
                        btnUnPick.Visible = False
                    Case "CANCELLED"
                        'btnPost.Visible = False
                        btnPostChk.Visible = False
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
                    uiFun.load_dropdown(DO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT  where " &
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                        "and imp_code = '" & Session("IMP_CODE") & "' " &
                                        "ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(DO_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT WHERE " &
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                          "and imp_code = '" & Session("IMP_CODE") & "' " &
                                          "and PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("DO_PROJECT_NO").ToString) & "' ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , , , True)
                End If


                lbl_PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("DO_PROJECT_NO").ToString) & "' ")
                'CUS_NAME.Text = DB.getValueFromSQL("SELECT CUS_NAME FROM WMS_CUSTOMER WHERE CUS_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CUS_CODE").ToString) & "' ")

                uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " &
                                              "storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                              "and imp_code = '" & Session("IMP_CODE") & "' " &
                                              "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

                CUS_CODE.SelectedValue = dt.Rows(0).Item("CUS_CODE").ToString

                uiFun.load_dropdown(DO_TRANS_TYPE, "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'WMS_DELV_ORDER.DO_TRANS_TYPE' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , , dt.Rows(0).Item("DO_TRANS_TYPE").ToString, True)

                SQLString = "SELECT IMP_CODE, STORER_CODE, DO_CODE, DOV_SEQ, DOV_VIDEO_NAME, DOV_WEIGHT, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, DOV_BOX_NO " &
                            " FROM WMS_DELV_ORDER_VIDEO " &
                            "WHERE WMS_DELV_ORDER_VIDEO.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                            "AND WMS_DELV_ORDER_VIDEO.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                            "AND WMS_DELV_ORDER_VIDEO.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "


                dtV = gDB.getDataTable(SQLString)

                'If dtV.Rows.Count > 0 Then
                '    GridView2.DataSource = dtV
                'Else
                '    GridView2.DataSource = Nothing
                'End If

                'GridView2.DataBind()


                SQLString = "SELECT IMP_CODE, STORER_CODE, DO_CODE, DOB_SEQ, DOB_BOX_NO, DOB_FTRACK_NO, DOB_SUB_FTRACK_NO, DOB_STATUS, DOB_DATE, DOB_RMKS " &
                            "FROM WMS_DO_BOX_TRACK " &
                            "WHERE WMS_DO_BOX_TRACK.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                            "AND WMS_DO_BOX_TRACK.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                            "AND WMS_DO_BOX_TRACK.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

                dtT = gDB.getDataTable(SQLString)

                'If dtT.Rows.Count > 0 Then
                '    GridView3.DataSource = dtT
                'Else
                '    GridView3.DataSource = Nothing
                'End If
                'GridView3.DataBind()

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
        SQLString += "pd.pad_origin, pd.sys_cb, pd.sys_cd, pd.sys_lub, pd.sys_lud, 'U' as mFlag, 0 as itm_total_qty, " &
                        "case when Charindex('-', pd.PAD_CARTON_NO )> 0 then " &
                            "right('00000000000000000000' + replace(SUBSTRING(pd.PAD_CARTON_NO,1,CHARINDEX('-',pd.PAD_CARTON_NO)),'-',''),20) " &
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
        '"ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, l.ILOC_BAL_QTY, " &
        SQLString = "SELECT distinct d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY,0 as stock_qty,d.pld_ss_qty,0 as ILOC_BAL_QTY, " &
                        "d.PLD_WH, d.PLD_LOC, d.PLD_ORG_LOC, d.PLD_REMARK, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, " &
                        "d.PLD_BATCH_NO, d.PLD_FOI_QTY, d.PLD_SERIAL_NO, d.PLD_QTY2, i.ITM_SERIAL_NO_YN, i.itm_type, " &
                        "Convert(int, PLD_SPLIT_FR) as PLD_SPLIT_FR_INT, Convert(int, PLD_SEQ) AS PLD_SEQ_INT, " &
                        "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
                        "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
                        "Convert(varchar,ISNULL(d.pld_expiry_date, l.ILOC_EXPIRY_DATE), " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " &
                        "isnull(dod.dod_disp_seq, 9999) as dod_disp_seq, b.BN_CSMS_CODE, " &
                        "PLD_TO_DRUM_ID, PLD_TO_DRUM_CABLE_LIST, PLD_TO_DRUM_ILOC_SEQ,'' as from_drum_id,  " &
                        "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag " &
                    "from WMS_DO_PICKLIST_D d " &
                    "left outer join WMS_ITEM_LOC_BAL l " &
                        "on d.IMP_CODE = l.IMP_CODE " &
                        "and d.STORER_CODE = l.STORER_CODE " &
                        "and d.PLD_ITEM_NO = l.ITM_CODE " &
                        "and l.ILOC_WH = d.PLD_WH " &
                        "and d.PLD_PACK_KEY = l.PACK_KEY " &
                        "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
                        "and d.PLD_LOC = l.ILOC_LOC " &
                        "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
                    "left outer join WMS_DATE_CODE DC " &
                        "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " &
                        "and d.IMP_CODE = DC.IMP_CODE " &
                        "and d.STORER_CODE = DC.STORER_CODE " &
                    "left outer join WMS_ITEM i " &
                        "on d.IMP_CODE = i.IMP_CODE " &
                        "and d.STORER_CODE = i.STORER_CODE " &
                        "and d.PLD_ITEM_NO = i.ITM_CODE " &
                        "and d.PLD_PACK_KEY = i.PACK_KEY " &
                    "LEFT OUTER JOIN ( " &
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                            "from wms_do_picklist_d p, wms_delv_order d " &
                            "where d.imp_code = p.imp_code " &
                            "and d.storer_code = p.storer_code " &
                            "and d.do_code = p.do_code " &
                            "and d.do_status = 'PICKED' " &
                            "and d.do_code <> '" & gU.dbEncode(doCode) & "' " &
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                        "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
                        "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
                        "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
                        "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                        "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                        "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                        "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
                    "left outer join (" &
                            "select imp_code, storer_code, do_code, dod_itm_code, dod_pack_key, min(dod_disp_seq) as dod_disp_seq " &
                            "from WMS_DELV_ORDER_D " &
                            "group by imp_code, storer_code, do_code, dod_itm_code, dod_pack_key) DOD " &
                        "ON d.IMP_CODE = DOD.IMP_CODE " &
                        "AND d.STORER_CODE = DOD.STORER_CODE " &
                        "AND d.DO_CODE = DOD.DO_CODE " &
                        "AND d.PLD_ITEM_NO = DOD.dod_itm_code " &
                        "AND d.PLD_PACK_KEY = DOD.dod_pack_key " &
                    "left outer join wms_wh_bin b " &
                        "on d.PLD_LOC = b.LOC_KEY " &
                    "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                    "and d.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "


        'SQLString = SQLString & " order by  DC.DC_DATE_CODE, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, PLD_WH, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN "
        SQLString = SQLString & " order by isnull(dod.dod_disp_seq, 9999), i.ITM_SKU_NO, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, Convert(int, PLD_SPLIT_FR), Convert(int, PLD_SEQ) "
        REM **********************
        Session("_M_OB_DO_TMP_pl_dt") = gDB.getDataTable(SQLString)
        REM **********************

        Dim objPlt As PickListTable
        objPlt = New PickListTable(DO_CO_CODE.Text, doCode, storerCode, Session("IMP_CODE"))

        objPlt.updateTotalQty(Session("_M_OB_DO_TMP_pl_dt"))

        objPlt.updateHoldBal(Session("_M_OB_DO_TMP_pl_dt"))

        objPlt = Nothing

        SQLString = "select MAX(CAST(PLD_SEQ AS int)) from WMS_DO_PICKLIST_D " &
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text.Trim) & "' "

        Session("_M_OB_DO_TMP_pl_seq") = DB.getValueFromSQL(SQLString)

        If Session("_M_OB_DO_TMP_pl_seq") = "" Then
            Session("_M_OB_DO_TMP_pl_seq") = "0"
        End If

        Session("_M_OB_DO_TMP_pl_del_list") = ""

        REM Generate Data Table from Detail
        SQLString = "SELECT d.DOD_SEQ, d.DOD_DISP_SEQ, d.DOD_PALLET_NO, d.DOD_CARTON_NO, d.DOD_PACK_NO, d.DOD_ITM_CODE, d.DOD_PACK_KEY,d.DOD_RETURN_QTY, " &
                        "d.DOD_ITM_DESC, d.DOD_PACK_TYPE, d.DOD_QTY, d.DOD_UOM, d.DOD_CUT_YN, d.DOD_QTY2, d.DOD_UOM2, d.DOD_PCS_UOM, d.DOD_TOTPCS, d.DOD_TOT_WGT, d.DOD_TOT_CBM, " &
                        "d.DOD_REM, d.DOD_TICKET_NO, d.DOD_WH_CODE,d.DOD_FL_CODE,d.DOD_LOC_WH, d.DOD_VND_CODE, d.DOD_BATCH_NO, d.IMP_CODE, d.STORER_CODE, d.DO_CODE, i.ITM_SKU_NO, i.ITM_DESC, " &
                        "Convert(varchar,d.DOD_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_EXPIRY_DATE, " &
                        "Convert(varchar,d.DOD_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_MANU_DATE, " &
                        "d.DOD_SS_QTY,d.DOD_PK_QTY,d.DOD_MIN_PROD_DATE,d.DOD_MIN_SHELF_LIFE,d.DOD_LAST_LOT,d.DOD_MAX_LOT," &
                        "'U' as mFlag, 1 as label_qty, '' as dod_ref_no, d.DOD_TROLLEY_ID, i.ITM_SERIAL_NO_YN, i.ITM_TYPE " &
                    "from WMS_DELV_ORDER_D d " &
                    "Left outer join WMS_ITEM i " &
                    "ON d.IMP_CODE = i.IMP_CODE " &
                    "and d.STORER_CODE = i.STORER_CODE " &
                    "and d.DOD_ITM_CODE = i.ITM_CODE " &
                    "and d.DOD_PACK_KEY = i.PACK_KEY " &
                    "where d.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " &
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
        Dim cancelSql As String = "update wms_delv_order " &
                     "set do_status = 'CANCELLED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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

    Protected Sub saveStorerBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveStorerBtn.Click
        Dim cancelSql As String = "update wms_delv_order " &
                     "set DO_STORER_REM = " & gU.convdbNVCData(gU.dbEncode(DO_STORER_REM.Text.Trim)) & ", " &
                     "DO_STORER_LUD = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and do_code = '" & gU.dbEncode(DO_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            uiFun.displayMsgNew(updtPnlAlert, "", "Record has been saved", Session("gLang"))

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

    Private Function SaveReserve(ByVal CUS_CODE As String, ByVal ITEM_CODE As String, ByVal LOT_NO As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime) As Int16
        Dim gConn As SqlConnection
        gConn = gDB.getConnection()

        'Dim transaction As SqlTransaction
        'transaction = gConn.BeginTransaction()

        Dim sql_string As String = ""
        sql_string += "insert into WMS_WAVEPICK_RSVD (imp_code, storer_code, CUS_CODE,ITEM_CODE,LOT_NO,QTY,ROUTE_ID,DO_DATE)"
        sql_string += "values ("
        sql_string += "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "',"
        sql_string += "'" & gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(CUS_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ITEM_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(LOT_NO.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(QTY.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ROUTE_ID.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(DO_DATE.ToString.Trim) & "'"
        sql_string += ")"

        Dim RowCount = gDB.amendData(sql_string, gConn, Nothing)
        Return RowCount
    End Function

    Private Function custRuleChk(ByVal ITEM_CODE As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime, ByVal Optional CUS_CODE As String = "", ByVal Optional LAST_LOT As String = "") As List(Of Dictionary(Of String, Object))
        Dim SQLString As String
        Dim dtRule As DataTable
        Dim dtStock As DataTable
        Dim returnList As New List(Of Dictionary(Of String, Object))

        SQLString = "Select a.*,IsNull(b.ITM_SHELF_LIFE,0)ITM_SHELF_LIFE from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE.Value) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and a.ILOC_EXPIRY_DATE > GetDate() and a.ILOC_BAL_QTY > 0 order by a.ILOC_EXPIRY_DATE "
        dtStock = gDB.getDataTable(SQLString)

        If dtStock IsNot Nothing And dtStock.Rows.Count > 0 Then
            If CUS_CODE <> "" Then
                CUS_CODE = CUS_CODE.Split(",")(0)
                SQLString = "Select MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,LOTS_CANNOT_BE_EARLIER from WMS_CUSTOMER_RULE where item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
                dtRule = gDB.getDataTable(SQLString)
                If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
                    SQLString = "Select MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,LOTS_CANNOT_BE_EARLIER from WMS_CUSTOMER Where  IMP_CODE='" & gU.dbEncode(IMP_CODE.Value) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
                    dtRule = gDB.getDataTable(SQLString)
                End If

                If dtRule.Rows.Count > 0 Then
                    Dim dtToday As DateTime = System.DateTime.Now
                    Dim PrevLot As DateTime = Nothing
                    returnList.Clear()
                    For Each rowStk In dtStock.Rows
                        Dim dtLot As DateTime = Convert.ToDateTime(rowStk("ILOC_EXPIRY_DATE").ToString)
                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE.Value) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "' and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "'"
                        Dim dtRSVD As DataTable = gDB.getDataTable(SQLString)
                        If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then

                            If (dtToday - dtLot).Days <= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                If Convert.ToInt16(rowStk("ITM_SHELF_LIFE")) <= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                    Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                                    If balQty > 0 Then
                                        Dim dictParam As New Dictionary(Of String, Object)()
                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                        dictParam.Add("DOD_MAX_LOT", "")
                                        If QTY <= balQty Then

                                            dictParam.Add("DOD_PK_QTY", QTY)
                                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                        Else
                                            dictParam.Add("DOD_PK_QTY", balQty)
                                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                            QTY = QTY - balQty
                                        End If
                                        returnList.Add(dictParam)
                                    End If
                                    'PrevLot = dtLot
                                End If
                            End If
                        ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then

                            If dtToday.AddDays((dtToday - dtLot).Days) <= Convert.ToDateTime(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                If dtLot.AddDays(Convert.ToInt16(rowStk("ITM_SHELF_LIFE"))) <= Convert.ToDateTime(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                    Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                                    If balQty > 0 Then
                                        Dim dictParam As New Dictionary(Of String, Object)()
                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                        dictParam.Add("DOD_MAX_LOT", "")
                                        If QTY <= balQty Then
                                            dictParam.Add("DOD_PK_QTY", QTY)
                                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                        Else

                                            dictParam.Add("DOD_PK_QTY", balQty)
                                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                            QTY = QTY - balQty
                                        End If
                                        returnList.Add(dictParam)
                                    End If
                                    'PrevLot = dtLot
                                End If
                            End If
                        ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then
                            If LAST_LOT IsNot Nothing Then
                                Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                If dtLot.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                    If dtLot > Last_LOT_DT Then
                                        Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                                        If balQty > 0 Then
                                            Dim dictParam As New Dictionary(Of String, Object)()
                                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                            dictParam.Add("DOD_MAX_LOT", "")
                                            If QTY <= balQty Then
                                                dictParam.Add("DOD_PK_QTY", QTY)
                                                Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                            Else

                                                dictParam.Add("DOD_PK_QTY", balQty)
                                                Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                                QTY = QTY - balQty
                                            End If
                                            returnList.Add(dictParam)

                                        End If
                                        'PrevLot = dtLot
                                    End If
                                End If
                            End If
                        ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                            If dtLot > Last_LOT_DT Then
                                Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                                If balQty > 0 Then
                                    Dim dictParam As New Dictionary(Of String, Object)()
                                    dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                    dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                    dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                    dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                    dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                    dictParam.Add("DOD_MAX_LOT", "")
                                    If QTY <= balQty Then
                                        dictParam.Add("DOD_PK_QTY", QTY)
                                        Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                    Else

                                        dictParam.Add("DOD_PK_QTY", balQty)
                                        Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                        QTY = QTY - balQty
                                    End If
                                    returnList.Add(dictParam)

                                End If
                                'PrevLot = dtLot
                            End If

                        ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" Then
                            If dtLot.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                                If balQty > 0 Then
                                    Dim dictParam As New Dictionary(Of String, Object)()
                                    dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                    dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                    dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                    dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                    dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                    dictParam.Add("DOD_MIN_PROD_DATE", If(dtRule.Rows(0)("MIN_PROD_LIFE") IsNot "", dtRule.Rows(0)("MIN_PROD_LIFE"), "0"))
                                    dictParam.Add("DOD_MIN_SHELF_LIFE", If(dtRule.Rows(0)("MIN_SELF_LIFE") IsNot "", dtRule.Rows(0)("MIN_SELF_LIFE"), "0"))
                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                    dictParam.Add("DOD_MAX_LOT", "")
                                    If QTY <= balQty Then

                                        dictParam.Add("DOD_PK_QTY", QTY)
                                        Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                    Else
                                        dictParam.Add("DOD_PK_QTY", balQty)
                                        Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                        QTY = QTY - balQty
                                    End If
                                    returnList.Add(dictParam)
                                End If
                                'PrevLot = dtLot
                            End If

                        Else


                            Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                            If balQty > 0 Then
                                Dim dictParam As New Dictionary(Of String, Object)()
                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                dictParam.Add("ILOC_BAL_QTY", 0)
                                dictParam.Add("DOD_MIN_PROD_DATE", DBNull.Value)
                                dictParam.Add("DOD_MIN_SHELF_LIFE", DBNull.Value)
                                dictParam.Add("DOD_LAST_LOT", "")
                                dictParam.Add("DOD_MAX_LOT", "")
                                If QTY <= balQty Then

                                    dictParam.Add("DOD_PK_QTY", QTY)
                                    Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                                Else

                                    dictParam.Add("DOD_PK_QTY", balQty)
                                    Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                                    QTY = QTY - balQty
                                End If
                                returnList.Add(dictParam)
                            End If

                        End If
                    Next
                End If
            Else
                For Each rowStk In dtStock.Rows
                    SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE.Value) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "' and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "'"
                    Dim dtRSVD As DataTable = gDB.getDataTable(SQLString)
                    Dim balQty = Double.Parse(rowStk("ILOC_BAL_QTY")) - Double.Parse(dtRSVD.Rows(0)("QTY"))
                    If balQty > 0 Then
                        Dim dictParam As New Dictionary(Of String, Object)()
                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                        dictParam.Add("ILOC_BAL_QTY", 0)
                        dictParam.Add("DOD_MIN_PROD_DATE", "0")
                        dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                        dictParam.Add("DOD_LAST_LOT", "")
                        dictParam.Add("DOD_MAX_LOT", "")
                        If QTY <= balQty Then

                            dictParam.Add("DOD_PK_QTY", QTY)
                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), QTY, ROUTE_ID, DO_DATE)

                        Else
                            dictParam.Add("DOD_PK_QTY", balQty)
                            Dim result = SaveReserve(CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE)
                            QTY = QTY - balQty
                        End If
                        returnList.Add(dictParam)
                    End If
                Next
            End If

        Else
            Dim dictParam As New Dictionary(Of String, Object)()
            dictParam.Add("ILOC_EXPIRY_DATE", "")
            dictParam.Add("ITM_CODE", ITEM_CODE)
            dictParam.Add("ILOC_BATCH_NO", "")
            dictParam.Add("ILOC_LOC", "")
            dictParam.Add("ILOC_WH", "")
            dictParam.Add("ILOC_BAL_QTY", 0)
            dictParam.Add("DOD_MIN_PROD_DATE", "0")
            dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
            dictParam.Add("DOD_LAST_LOT", "")
            dictParam.Add("DOD_MAX_LOT", "")
            dictParam.Add("DOD_PK_QTY", "0")
            returnList.Add(dictParam)
        End If
        Return returnList
    End Function

    Private Function IsRule(ByVal CUS_CODE As String, ByVal ITEM_CODE As String) As DataRow
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        SQLString = "Select * from WMS_CUSTOMER_RULE where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
        dtRule = gDB.getDataTable(SQLString)
        If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
            SQLString = "Select MPL_FLAG,MIN_PROD_LIFE,MSL_FLAG,MIN_SELF_LIFE,MB_FLAG,MAX_BATCHES,LOTS_CANNOT_BE_EARLIER from WMS_CUSTOMER Where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and IMP_CODE='" & gU.dbEncode(IMP_CODE.Value) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
            dtRule = gDB.getDataTable(SQLString)
        End If
        dtRule.Columns.Add("RuleCount")
        If dtRule IsNot Nothing AndAlso dtRule.Rows.Count > 0 Then
            If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 4

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            End If
        Else
            Return Nothing
        End If
        Return dtRule.Rows(0)
    End Function

    'Protected Sub btnPicking_Click(sender As Object, e As System.EventArgs) Handles btnPicking.Click
    '    Dim SQLString As String
    '    Dim updateSQL As String = ""
    '    Dim nextNo As String
    '    Dim coHdrDATES As DataTable
    '    Dim coHdrGrp As DataTable
    '    Dim coDtl As DataTable
    '    Dim Qty As Double = 0
    '    Dim gConn = gDB.getConnection()
    '    Dim SeqNo As Int16 = 0

    '    Try
    '        SQLString = "Select Distinct CO_DATE FROM WMS_CUST_ORDER SS Where ss.ROUTE_ID != '' and ss.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "'"
    '        coHdrDATES = gDB.getDataTable(SQLString)
    '        For Each dtRow As DataRow In coHdrDATES.Rows
    '            SQLString = "SELECT Distinct SS.ROUTE_ID,STUFF((SELECT ',' + US.CO_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "' FOR XML PATH('')), 1, 1, '') [CO_CODES],STUFF((SELECT ',' + US.CUS_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "' FOR XML PATH('')), 1, 1, '') [CUS_CODES],STUFF((SELECT ',' + US.CUS_NAME FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "' FOR XML PATH('')), 1, 1, '') [CUS_NAMES] FROM WMS_CUST_ORDER SS Where ss.ROUTE_ID != '' and ss.CO_Date ='" & gU.dbEncode(dtRow("CO_DATE")) & "' and ss.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "'"
    '            coHdrGrp = gDB.getDataTable(SQLString)
    '            For Each dtGrpRow As DataRow In coHdrGrp.Rows
    '                nextNo = DB.getDocNo("Do", gConn, Nothing)
    '                SQLString = "INSERT INTO WMS_DELV_ORDER(IMP_CODE ,STORER_CODE ,DO_CODE,ROUTE_ID ,DO_STATUS ,DO_ISSUED_BY ,DO_CO_CODE ,DO_CUS_REF_NO ,DO_DATE ,CUS_CODE ,CUS_NAME ,DO_POSTED_DATE ,DO_POSTED_BY ,DO_TYPE ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
    '                SQLString += "VALUES("
    '                SQLString += "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "',"
    '                SQLString += "'" & gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) & "',"
    '                SQLString += "'" & nextNo & "',"
    '                SQLString += "'" & gU.dbEncode(dtGrpRow("ROUTE_ID").ToString.Trim) & "',"
    '                SQLString += "'DRAFT',"
    '                SQLString += "'" & Session("usr_id") & "',"
    '                SQLString += "'" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & "',"
    '                SQLString += "'NA',"
    '                SQLString += "'" & dtRow("CO_DATE").ToString.Trim & "',"
    '                SQLString += "'" & gU.dbEncode(dtGrpRow("CUS_CODES").ToString.Trim) & "',"
    '                SQLString += "N'" & gU.dbEncode(dtGrpRow("CUS_NAMES").ToString.Trim) & "'," '
    '                SQLString += "GETDATE(),"
    '                SQLString += "'" & Session("usr_id") & "',"
    '                SQLString += "'NA',"
    '                SQLString += "'" & Session("usr_id") & "',"
    '                SQLString += "GETDATE(),"
    '                SQLString += "GETDATE(),"
    '                SQLString += "'" & Session("usr_id") & "'"
    '                SQLString += ")"
    '                gDB.amendData(SQLString, gConn, Nothing)

    '                SQLString = "Select a.*,b.ROUTE_ID,b.CUS_CODE,b.CUS_NAME from WMS_CUST_ORDER_D a Inner join WMS_CUST_ORDER b On a.CO_CODE = b.CO_CODE " &
    '                "WHERE  " &
    '                "b.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "'" &
    '                "and b.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "' and a.CO_CODE in (" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & ")  Order by b.ROUTE_ID"
    '                coDtl = gDB.getDataTable(SQLString)
    '                coDtl.Columns.Add("IsRule")
    '                coDtl.Columns.Add("MIN_SELF_LIFE")
    '                coDtl.Columns.Add("MIN_PROD_LIFE")
    '                coDtl.Columns.Add("MAX_BATCHES")
    '                coDtl.Columns.Add("MPL_FLAG")
    '                coDtl.Columns.Add("MB_FLAG")
    '                coDtl.Columns.Add("MSL_FLAG")
    '                coDtl.Columns.Add("LOTS_CANNOT_BE_EARLIER")

    '                'Dim dtView As New DataView(codt)
    '                'Dim dtTemp As DataTable = dtView.ToTable(True, "CUS_CODE", "COD_ITM_CODE")
    '                Dim dtWPickNORULE As New DataTable
    '                Dim dtWPick1RULE As New DataTable
    '                Dim dtWPickRULE As New DataTable
    '                Dim grpQTY As Double = 0
    '                Dim grpCus As String = ""
    '                Dim custCodes As New List(Of String)
    '                Dim prevItem As String = coDtl.Rows(0)("COD_ITM_CODE").ToString()
    '                Dim RowList As New List(Of DataRow)
    '                dtWPickNORULE = coDtl.Clone()
    '                dtWPick1RULE = coDtl.Clone()
    '                dtWPickRULE = coDtl.Clone()
    '                dtWPickNORULE.Rows.Clear()
    '                dtWPick1RULE.Rows.Clear()
    '                dtWPickRULE.Rows.Clear()

    '                For Each row As DataRow In coDtl.Rows
    '                    Dim RuleRow = IsRule(row("CUS_CODE").ToString(), row("COD_ITM_CODE").ToString())
    '                    If RuleRow IsNot Nothing Then
    '                        row("IsRule") = RuleRow("RuleCount")
    '                        row("MIN_SELF_LIFE") = RuleRow("MIN_SELF_LIFE")
    '                        row("MIN_PROD_LIFE") = RuleRow("MIN_PROD_LIFE")
    '                        row("MAX_BATCHES") = RuleRow("MAX_BATCHES")

    '                        row("MPL_FLAG") = RuleRow("MPL_FLAG")
    '                        row("MSL_FLAG") = RuleRow("MSL_FLAG")
    '                        row("MB_FLAG") = RuleRow("MB_FLAG")
    '                        row("LOTS_CANNOT_BE_EARLIER") = RuleRow("LOTS_CANNOT_BE_EARLIER")

    '                    Else
    '                        row("IsRule") = "0"
    '                        row("MIN_SELF_LIFE") = "0"
    '                        row("MIN_PROD_LIFE") = "0"
    '                        row("MAX_BATCHES") = "0"
    '                        row("MPL_FLAG") = "0"
    '                        row("MSL_FLAG") = "0"
    '                        row("MB_FLAG") = "0"
    '                        row("LOTS_CANNOT_BE_EARLIER") = "0"
    '                    End If
    '                Next
    '                'Dim tmpTable = coDtl.DefaultView.ToTable(True, "CUS_CODE", "COD_ITM_CODE")

    '                'No Rule
    '                Dim tmpTable = If(coDtl.Select("IsRule = 0").Length > 0, coDtl.Select("IsRule = 0").CopyToDataTable(), New DataTable())
    '                If tmpTable.Rows.Count > 0 Then
    '                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE")
    '                    For Each row As DataRow In tmpTableDistinct.Rows


    '                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "'")
    '                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
    '                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "'")
    '                        If tmpArra.Count > 0 Then
    '                            For Each item As DataRow In tmpArra
    '                                custCodes.Add(item("CUS_CODE"))
    '                            Next
    '                            Dim rgRow = coDtl.NewRow()
    '                            'rgRow.ItemArray = row.ItemArray
    '                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
    '                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
    '                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
    '                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
    '                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

    '                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
    '                            rgRow("COD_PCS_UOM") = tmpArra(0)("COD_PCS_UOM").ToString()
    '                            rgRow("COD_TOTPCS") = tmpArra(0)("COD_TOTPCS").ToString()
    '                            rgRow("COD_TOT_WGT") = tmpArra(0)("COD_TOT_WGT").ToString()
    '                            rgRow("COD_TOT_CBM") = tmpArra(0)("COD_TOT_CBM").ToString()

    '                            rgRow("COD_QTY") = grpQTY
    '                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
    '                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
    '                            dtWPickNORULE.Rows.Add(rgRow.ItemArray)
    '                            RowList.Add(row)
    '                        End If


    '                    Next
    '                End If

    '                tmpTable = If(coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND LOTS_CANNOT_BE_EARLIER = 0").Length > 0, coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND LOTS_CANNOT_BE_EARLIER = 0").CopyToDataTable(), New DataTable())
    '                If tmpTable.Rows.Count > 0 Then
    '                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "MIN_SELF_LIFE")
    '                    For Each row As DataRow In tmpTableDistinct.Rows

    '                        custCodes.Clear()
    '                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "'")
    '                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
    '                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "'")
    '                        If tmpArra.Count > 0 Then
    '                            For Each item As DataRow In tmpArra
    '                                custCodes.Add(item("CUS_CODE"))
    '                            Next
    '                            Dim rgRow = coDtl.NewRow()
    '                            'rgRow.ItemArray = row.ItemArray
    '                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
    '                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
    '                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
    '                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
    '                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

    '                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
    '                            rgRow("COD_PCS_UOM") = tmpArra(0)("COD_PCS_UOM").ToString()
    '                            rgRow("COD_TOTPCS") = tmpArra(0)("COD_TOTPCS").ToString()
    '                            rgRow("COD_TOT_WGT") = tmpArra(0)("COD_TOT_WGT").ToString()
    '                            rgRow("COD_TOT_CBM") = tmpArra(0)("COD_TOT_CBM").ToString()

    '                            rgRow("COD_QTY") = grpQTY
    '                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
    '                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())

    '                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
    '                            RowList.Add(row)

    '                        End If


    '                    Next

    '                End If

    '                '
    '                tmpTable = If(coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0").Length > 0, coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0").CopyToDataTable(), New DataTable())
    '                If tmpTable.Rows.Count > 0 Then
    '                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_PALLET_NO")
    '                    For Each row As DataRow In tmpTableDistinct.Rows

    '                        custCodes.Clear()
    '                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "'")
    '                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
    '                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "'")
    '                        If tmpArra.Count > 0 Then
    '                            For Each item As DataRow In tmpArra
    '                                custCodes.Add(item("CUS_CODE"))
    '                            Next
    '                            Dim rgRow = coDtl.NewRow()
    '                            'rgRow.ItemArray = row.ItemArray
    '                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
    '                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
    '                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
    '                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
    '                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

    '                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
    '                            rgRow("COD_PCS_UOM") = tmpArra(0)("COD_PCS_UOM").ToString()
    '                            rgRow("COD_TOTPCS") = tmpArra(0)("COD_TOTPCS").ToString()
    '                            rgRow("COD_TOT_WGT") = tmpArra(0)("COD_TOT_WGT").ToString()
    '                            rgRow("COD_TOT_CBM") = tmpArra(0)("COD_TOT_CBM").ToString()

    '                            rgRow("COD_QTY") = grpQTY
    '                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
    '                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())

    '                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
    '                            RowList.Add(row)
    '                        End If


    '                    Next
    '                End If
    '                '
    '                SeqNo = 1
    '                For Each rowD As DataRow In dtWPickNORULE.Rows
    '                    Dim LOTNO As String = ""
    '                    Dim LOTQTY As Double = 0
    '                    Dim chkdList = custRuleChk(rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), "")
    '                    For Each itemDict In chkdList
    '                        SQLString = ""

    '                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
    '                        SQLString += "VALUES("
    '                        SQLString += "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) & "',"
    '                        SQLString += "'" & nextNo & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_WH").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
    '                        SQLString += "'" & (SeqNo) & "',"
    '                        SQLString += "0,"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
    '                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
    '                        SQLString += "'0',"
    '                        SQLString += "'0',"
    '                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
    '                        SQLString += "0,"
    '                        SQLString += "0,"
    '                        SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
    '                        SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & ","
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
    '                        SQLString += "'" & Session("usr_id") & "',"
    '                        SQLString += "GETDATE(),"
    '                        SQLString += "GETDATE(),"
    '                        SQLString += "'" & Session("usr_id") & "'"
    '                        SQLString += ")"
    '                        SeqNo = SeqNo + 1
    '                        gDB.amendData(SQLString, gConn, Nothing)
    '                    Next
    '                Next

    '                For Each rowD As DataRow In dtWPick1RULE.Rows
    '                    Dim LOTNO As String = ""
    '                    Dim LOTQTY As Double = 0
    '                    Dim chkdList = custRuleChk(rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
    '                    For Each itemDict In chkdList
    '                        SQLString = ""

    '                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
    '                        SQLString += "VALUES("
    '                        SQLString += "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) & "',"
    '                        SQLString += "'" & nextNo & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_WH").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
    '                        SQLString += "'" & (SeqNo) & "',"
    '                        SQLString += "0,"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
    '                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
    '                        SQLString += "'0',"
    '                        SQLString += "'0',"
    '                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
    '                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
    '                        SQLString += "0,"
    '                        SQLString += "0,"
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
    '                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
    '                        SQLString += "'" & Session("usr_id") & "',"
    '                        SQLString += "GETDATE(),"
    '                        SQLString += "GETDATE(),"
    '                        SQLString += "'" & Session("usr_id") & "'"
    '                        SQLString += ")"
    '                        SeqNo = SeqNo + 1
    '                        gDB.amendData(SQLString, gConn, Nothing)
    '                    Next
    '                Next
    '                tmpTable.Clear()
    '                tmpTable = If(coDtl.Select("IsRule > 1").Length > 0, coDtl.Select("IsRule > 1").CopyToDataTable(), New DataTable())
    '                If tmpTable.Rows.Count > 0 Then

    '                    'For Each row As DataRow In tmpTable.Rows


    '                    For Each rowD As DataRow In tmpTable.Rows
    '                        Dim LOTNO As String = ""
    '                        Dim LOTQTY As Double = 0
    '                        Dim chkdList = custRuleChk(rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
    '                        For Each itemDict In chkdList
    '                            SQLString = ""

    '                            SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
    '                            SQLString += "VALUES("
    '                            SQLString += "'" & gU.dbEncode(IMP_CODE.Value.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) & "',"
    '                            SQLString += "'" & nextNo & "',"
    '                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_WH").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
    '                            SQLString += "'" & (SeqNo) & "',"
    '                            SQLString += "0,"
    '                            SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
    '                            SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
    '                            SQLString += "'0',"
    '                            SQLString += "'0',"
    '                            SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
    '                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
    '                            SQLString += "" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & ","
    '                            SQLString += "" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & ","
    '                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & ","
    '                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & ","
    '                            SQLString += "0,"
    '                            SQLString += "0,"
    '                            SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
    '                            SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & ","
    '                            SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
    '                            SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
    '                            SQLString += "'" & Session("usr_id") & "',"
    '                            SQLString += "GETDATE(),"
    '                            SQLString += "GETDATE(),"
    '                            SQLString += "'" & Session("usr_id") & "'"  'dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
    '                            SQLString += ")"
    '                            SeqNo = SeqNo + 1
    '                            gDB.amendData(SQLString, gConn, Nothing)
    '                        Next
    '                    Next

    '                    'Next


    '                End If

    '                SQLString = "Select * from WMS_DELV_ORDER_D " &
    '                "WHERE  " &
    '                "DO_CODE = '" & gU.dbEncode(nextNo) & "'"
    '                Dim emptyHdr = gDB.getDataTable(SQLString)
    '                If emptyHdr Is Nothing Or emptyHdr.Rows.Count <= 0 Then
    '                    SQLString = "Delete from WMS_DELV_ORDER " &
    '                "WHERE  " &
    '                "DO_CODE = '" & gU.dbEncode(nextNo) & "'"
    '                    gDB.amendData(SQLString, gConn, Nothing)

    '                End If
    '                reloadPage("Wave Pick Success!")
    '            Next
    '        Next
    '    Catch ex As Exception
    '        reloadPage(ex.Message)
    '    End Try
    'End Sub

    'Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
    '    Dim updateSQL As String = ""
    '    Dim gconn As SqlConnection
    '    Dim transaction As SqlTransaction

    '    gconn = gDB.getConnection()
    '    transaction = gConn.BeginTransaction()
    '    Dim doCode, storerCode As String
    '    Try
    '        doCode = DO_CODE.Text
    '        storerCode = STORER_CODE.SelectedValue
    '        updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='RELEASED'" &
    '                        " Where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "' and DO_STATUS='DRAFT'"
    '        gDB.amendData(updateSQL)

    '        updateSQL = "Update WMS_CUST_ORDER set CO_STATUS='PICKED'" &
    '                        " Where imp_code ='" & gU.dbEncode(Session("imp_code")) & "' and storer_code ='" & gU.dbEncode(storerCode) & "' and CO_STATUS='NEW'"
    '        gDB.amendData(updateSQL)

    '        'transaction.Commit()

    '        reloadPage("The order has been released!!")

    '    Catch ex As Exception
    '        If Not transaction Is Nothing Then
    '            transaction.Rollback()
    '            transaction = Nothing
    '        End If

    '        If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
    '            Response.Write(ex.Message)
    '            uiFun.displayMsg(Me, "1008", "", Session("gLang"))
    '        Else
    '            uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
    '        End If

    '    Finally
    '        If gconn IsNot Nothing Then
    '            If gconn.State = ConnectionState.Open Then
    '                gconn.Close()
    '                gconn.Dispose()
    '            End If
    '        End If
    '    End Try
    'End Sub

    'Protected Sub btnPickingDelete_Click(sender As Object, e As System.EventArgs) Handles btnPickingDelete.Click
    '    Dim updateSQL As String = ""
    '    Dim SQLstring As String = ""
    '    Dim codt As DataTable
    '    Dim codtl As DataTable
    '    Dim gconn As SqlConnection
    '    Dim transaction As SqlTransaction
    '    gconn = gDB.getConnection()
    '    transaction = gConn.BeginTransaction()
    '    Dim doCode, storerCode As String
    '    Try
    '        doCode = DO_CODE.Text
    '        storerCode = STORER_CODE.SelectedValue
    '        SQLstring = "SELECT DO_CODE FROM WMS_DELV_ORDER " &
    '        " WHERE DO_STATUS ='DRAFT'"
    '        codt = gDB.getDataTable(SQLstring)
    '        For Each rowD As DataRow In codt.Rows
    '            updateSQL = "Delete from WMS_DELV_ORDER_D " &
    '                            "Where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "' and DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "'"
    '            gDB.amendData(updateSQL)

    '            updateSQL = "Delete from WMS_DELV_ORDER " &
    '                            " Where DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "' and imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "'"
    '            gDB.amendData(updateSQL)
    '        Next

    '        SQLstring = "SELECT CO_CODE FROM WMS_CUST_ORDER " &
    '        " WHERE CO_STATUS ='NEW'"
    '        codtl = gDB.getDataTable(SQLstring)
    '        For Each rowD As DataRow In codtl.Rows
    '            updateSQL = "Delete from WMS_CUST_ORDER_D " &
    '                            " Where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "' and CO_CODE='" & (rowD("CO_CODE").ToString.Trim) & "'"
    '            gDB.amendData(updateSQL)

    '            updateSQL = "Delete from WMS_CUST_ORDER " &
    '                            " Where CO_CODE='" & (rowD("CO_CODE").ToString.Trim) & "' and imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(storerCode) & "'"
    '            gDB.amendData(updateSQL)
    '        Next
    '        reloadPage("The order has been deleted!")
    '    Catch ex As Exception
    '        If Not transaction Is Nothing Then
    '            transaction.Rollback()
    '            transaction = Nothing
    '        End If

    '        If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
    '            Response.Write(ex.Message)
    '            uiFun.displayMsg(Me, "1008", "", Session("gLang"))
    '        Else
    '            uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
    '        End If

    '    Finally
    '        If gconn IsNot Nothing Then
    '            If gconn.State = ConnectionState.Open Then
    '                gconn.Close()
    '                gconn.Dispose()
    '            End If
    '        End If
    '    End Try
    'End Sub

    Private Function hasSerialItem() As Boolean
        Dim pl_dt As DataTable
        Dim i As Integer

        pl_dt = Session("_M_OB_DO_TMP_pl_dt")

        For i = 0 To pl_dt.Rows.Count - 1
            If pl_dt.Rows(i).Item("itm_serial_no_yn").ToString.Trim = "Y" AndAlso pl_dt.Rows(i).Item("itm_type").ToString.Trim <> "CABLE" Then
                Return True
            End If
        Next

        Return False
    End Function

    Protected Sub btnPostChk_Click(sender As Object, e As System.EventArgs) Handles btnPostChk.Click
        If Not prePostChk() Then
            Exit Sub
        End If

        If hasSerialItem() Then
            'ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "POST_CHECK", "postChk();", True)
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "POST_CHECK", "postChkReload();", True)
            'postChkReload
        Else
            postChkOK(True)
        End If
    End Sub

    Private Sub postChkOK(ByVal isPreChecked As Boolean)
        Dim updateSql, selectSql As String
        Dim pl_dt As DataTable
        Dim itmKey As String
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim gConn As SqlConnection
        Dim transaction As SqlTransaction
        Dim lCOD_SEQ As String
        Dim errorMsg As String = ""

        If Not isPreChecked Then
            If Not prePostChk() Then
                Exit Sub
            End If
        End If

        'uiFun.displayMsgNew(updtPnlAlert, "", "Test message!", Session("gLang"))

        'Exit Sub

        gConn = gDB.getConnection()

        transaction = gConn.BeginTransaction()

        Try
            qtyDict = New Dictionary(Of String, Double)
            cbmDict = New Dictionary(Of String, Double)
            wgtDict = New Dictionary(Of String, Double)

            For Each rows As DataRow In dt.Rows
                itmKey = gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" &
                        gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" &
                        gU.decodeNull(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" &
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
                            uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                              "don\'t have enough total stock to delivery!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                              "沒有足夠總貨存出貨!", Session("gLang"))
                        End If
                        transaction.Rollback()
                        gConn.Close()
                        gConn.Dispose()
                        Exit Sub
                    End If
                End If
            Next

            Dim haveReDrumYN As String = "N"
            Dim hasCable As Boolean = False
            'Check have redrum
            For Each rows As DataRow In pl_dt.Rows
                If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
                    haveReDrumYN = "Y"
                End If

                If rows.Item("ITM_TYPE").ToString.Trim = "CABLE" Then
                    hasCable = True
                End If
            Next

            If hasCable Then
                If haveReDrumYN = "N" And 1 = 2 Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Re-Drum should be performed first!", Session("gLang"))
                    transaction.Rollback()
                    gConn.Close()
                    gConn.Dispose()
                    Exit Sub
                End If
            End If

            'Check for each loc bal
            For Each rows As DataRow In pl_dt.Rows
                Dim SrchStr As String = ""
                SrchStr += "select ILOC_BAL_QTY,ILOC_BATCH_NO,ILOC_LOC,ILOC_PALLET_NO, Convert(varchar, ILOC_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as ExpDt, Convert(varchar, ILOC_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as ManuDt, Convert(datetime, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' "
                'SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, ""), "000")) & "' "

                If rows.Item("pld_batch_no").ToString.Trim <> "" Then
                    SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' "
                Else
                    SrchStr += "AND ISNULL(ILOC_BATCH_NO,'') = '' "
                End If

                Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)


                MANU_DATE = qtydt.Rows(0).Item("ManuDt").ToString.Trim 'rows.Item("PLD_MANU_DATE").ToString.Trim
                EXP_DATE = qtydt.Rows(0).Item("ExpDt").ToString.Trim 'rows.Item("PLD_EXPIRY_DATE").ToString.Trim

                If qtydt.Rows.Count > 0 Then
                    Dim stQTY As Double = gU.decodeEmptyCdbl(qtydt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
                    Dim plQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)
                    Dim ssQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_ss_qty").ToString.Trim, 0)

                    If ssQTY > 0 Then
                        SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                        Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                        If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                            Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                            Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('" + gU.dbEncode(IMP_CODE.Value.Trim) + "','" + gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                            gDB.amendData(insertSql, gConn, transaction)
                        End If
                    End If

                    If stQTY < plQTY Then

                        Dim qtyTmpdt As DataTable = New DataTable()

                        SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                        SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                        SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                        SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                        SrchStr += "AND ILOC_BAL_QTY >= '" & (plQTY - stQTY) & "' "
                        SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                        SrchStr += "AND ILOC_LOC != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                        SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                        qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)

                        If qtyTmpdt Is Nothing OrElse qtyTmpdt.Rows.Count = 0 Then
                            SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                            SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                            SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                            SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                            SrchStr += "AND ILOC_BAL_QTY >= '" & (plQTY - stQTY) & "' "
                            SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                            SrchStr += "AND ILOC_LOC+ILOC_BATCH_NO != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) + gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                            qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)
                        End If
                        If qtyTmpdt IsNot Nothing AndAlso qtyTmpdt.Rows.Count > 0 AndAlso Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY")) >= (plQTY - stQTY) Then
                            Dim stNewQTY As Double = 0
                            stNewQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))

                            If (stNewQTY + stQTY) >= plQTY Then
                                INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, (plQTY - stQTY), qtydt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)

                            End If

                            'rows.Item("pld_batch_no") = qtyTmpdt.Rows(0)("ILOC_BATCH_NO")
                            'rows.Item("pld_area") = qtyTmpdt.Rows(0)("ILOC_AREA")
                            'rows.Item("pld_loc") = qtyTmpdt.Rows(0)("ILOC_LOC")

                            'MANU_DATE = qtyTmpdt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

                            'EXP_DATE = qtyTmpdt.Rows(0).Item("EXPIRY_DATE").ToString.Trim



                            SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                            Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                            If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                                Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                                Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('" + gU.dbEncode(IMP_CODE.Value.Trim) + "','" + gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                                gDB.amendData(insertSql, gConn, transaction)
                            End If
                        Else
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                              "don\'t have enough stock to delivery!", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                              gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                              "沒有足夠貨存出貨!", Session("gLang"))
                            End If

                            transaction.Rollback()
                            gConn.Close()
                            gConn.Dispose()
                            Exit Sub
                        End If
                    End If

                    If MANU_DATE = "" Then
                        MANU_DATE = qtydt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim
                    End If
                    If EXP_DATE = "" Then
                        EXP_DATE = qtydt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                    End If
                Else
                    Dim plQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)
                    SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_LOC,ILOC_PALLET_NO,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE,ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                    SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                    SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                    SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                    SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' "
                    SrchStr += "AND ILOC_LOC != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                    SrchStr += "AND ILOC_BAL_QTY >= '" & (plQTY) & "' "
                    SrchStr += "AND ILOC_LOC+ILOC_BATCH_NO != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) + gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                    Dim qtyTmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                    If qtyTmpdt IsNot Nothing AndAlso qtyTmpdt.Rows.Count > 0 AndAlso Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY")) > plQTY Then
                        Dim stQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))
                        If stQTY < plQTY Then

                            'rows.Item("pld_batch_no") = qtyTmpdt.Rows(0)("ILOC_BATCH_NO")
                            'rows.Item("pld_area") = qtyTmpdt.Rows(0)("ILOC_AREA")
                            'rows.Item("pld_loc") = qtyTmpdt.Rows(0)("ILOC_LOC")

                            'MANU_DATE = qtyTmpdt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

                            'EXP_DATE = qtyTmpdt.Rows(0).Item("EXPIRY_DATE").ToString.Trim

                            Dim stNewQTY As Double = 0
                            stNewQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))
                            If stNewQTY >= plQTY Then
                                INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, (plQTY - stQTY), qtydt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)

                            End If

                            SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                            Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                            If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                                Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                                Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('" + gU.dbEncode(IMP_CODE.Value.Trim) + "','" + gU.dbEncode(STORER_CODE.SelectedValue.ToString.Trim) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                                gDB.amendData(insertSql, gConn, transaction)
                            End If
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                          "don\'t have stock to delivery!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " &
                                          gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                          "沒有此貨存出貨!", Session("gLang"))
                        End If

                        transaction.Rollback()
                        gConn.Close()
                        gConn.Dispose()
                        Exit Sub
                    End If
                End If

                Dim Remark As String = "DO_CODE: " + DO_CODE.Text + ", ITEM_CODE: " + rows.Item("pld_item_no").ToString.Trim + ", PICKED_QTY: " + rows.Item("pld_item_qty").ToString.Trim
                Dim insertLog As String = "insert into [dbo].[ActionLogs] values ('" + gU.dbEncode(Remark) + "','DO_POST','WMS_DELV_ORDER',GetDate(),'" + Session("usr_id").ToString + "')"
                gDB.amendData(insertLog, gConn, transaction)

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


                itmKey = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" &
                         gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" &
                         gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" &
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
                    updateSql = "update WMS_CUST_ORDER_D " &
                                "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                    "sys_lud = Getdate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

                    gDB.amendData(updateSql, gConn, transaction)

                Else
                    selectSql = "select COD_SEQ as value " &
                                "from WMS_CUST_ORDER_D " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " &
                                "AND ISNULL(COD_POST_QTY, 0) < ISNULL(COD_QTY, 0) " &
                                "order by convert(int, COD_SEQ) "

                    lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

                    updateSql = "update WMS_CUST_ORDER_D " &
                                "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                    "sys_lud = Getdate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

                    gDB.amendData(updateSql, gConn, transaction)

                End If

                'Update coh_rel_qty for Hold stock logic
                updateSql = "update wms_cust_order_hold " &
                            "set COH_REL_QTY = case when ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) < ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " then wms_cust_order_hold.coh_in_stock_qty else ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " &
                                "sys_lub = '" & Session("usr_id") & "', " &
                                "sys_lud = Getdate() " &
                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                            "and wms_cust_order_hold.coh_status <> 'RELEASE' " &
                            "and ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) > 0 " &
                            "and ISNULL(wms_cust_order_hold.COH_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " &
                            "and ISNULL(wms_cust_order_hold.COH_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', '') " &
                            "AND wms_cust_order_hold.COH_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " &
                            "AND wms_cust_order_hold.COH_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "' " &
                            "and exists (" &
                                "select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                "where wms_cust_order_hold.imp_code = d.imp_code " &
                                "and wms_cust_order_hold.storer_code = d.storer_code " &
                                "and wms_cust_order_hold.co_code = d.co_code " &
                                "and wms_cust_order_hold.cod_seq = d.cod_seq " &
                                "and c.imp_code = d.imp_code " &
                                "and c.storer_code = d.storer_code " &
                                "and c.co_code = d.co_code " &
                                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED')) "

                gDB.amendData(updateSql, gConn, transaction)

                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                st.UpdateStockTrans("OUT", gConn, transaction)

                st.UpdateStockBalTrans("OUT", gConn, transaction)

                If rows.Item("pld_serial_no").ToString.Trim <> "" Then
                    st.UpdateStockSerialTrans("OUT", gConn, transaction)

                    st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                End If

                REM Update Transcation to PickList
                updateSql = "update WMS_DO_PICKLIST_D " &
                            "set PLD_TX_ID = '" & st.IO_SYS_SEQ & "', " &
                            "sys_lub = '" & Session("usr_id") & "', " &
                            "sys_lud = Getdate() " &
                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' " &
                            "and PLD_SEQ = '" & gU.dbEncode(rows.Item("PLD_SEQ").ToString.Trim) & "' "

                gDB.amendData(updateSql, gConn, transaction)

            Next

            For Each rows As DataRow In pl_dt.Rows
                If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
                    st.cableReDrum(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, gConn, transaction)
                End If
            Next


            updateSql = "update WMS_CUST_ORDER " &
                        "set CO_STATUS = 'CLOSED', " &
                            "sys_lub = '" & Session("usr_id") & "', " &
                            "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                        "AND NOT EXISTS (" &
                            "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                            "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                            "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                            "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                            "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                            "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

            gDB.amendData(updateSql, gConn, transaction)

            updateSql = "update WMS_CUST_ORDER " &
                        "set CO_STATUS = 'PARTIAL', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                        "AND EXISTS (" &
                            "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                            "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                            "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                            "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                            "AND D.COD_POST_QTY > 0) " &
                        "AND EXISTS (" &
                            "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                            "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                            "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                            "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                            "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                            "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

            gDB.amendData(updateSql, gConn, transaction)

            updateSql = "update WMS_DELV_ORDER " &
                        "set DO_STATUS = 'POSTED', DO_SAMPLE_CHECKED ='1', " &
                        "DO_POSTED_DATE = Getdate(), " &
                        "DO_POSTED_BY = '" & Session("usr_id") & "', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "

            gDB.amendData(updateSql, gConn, transaction)

            transaction.Commit()
            transaction = Nothing

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
            updateSql = "Delete from WMS_WAVEPICK_RSVD " &
                           " Where DO_CODE='" & DO_CODE.Text.Trim & "' and imp_code='" & gU.dbEncode(IMP_CODE.Value.ToString) & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue.ToString) & "'"
            gDB.amendData(updateSql)
            reloadPage("Record has been Posted successfully!")


        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
            'Throw ex

            uiFun.displayMsgNew(updtPnlAlert, "", "Post failed! " & ex.Message, Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        'If successFlag = True Then
        '    'Dim rmtPost As New RemotePost
        '    'rmtPost.Url = "DOMain.aspx"
        '    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
        '    'rmtPost.Add("DO_CODE", DO_CODE.Text)
        '    'rmtPost.alertMsg = "Record has been Posted successfully!"
        '    'rmtPost.Post()
        'End If

    End Sub

    Private Function prePostChk() As Boolean
        Dim SQLString As String
        Dim checkdt As DataTable

        SQLString = "SELECT CO_STATUS FROM WMS_CUST_ORDER M " &
                    "WHERE CO_STATUS = 'CLOSED' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' "

        checkdt = gDB.getDataTable(SQLString)

        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The SIR status is Closed, please re-open before post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The SIR status is Closed, please re-open before post DO!", Session("gLang"))
            End If
            Return False
        End If

        SQLString = "SELECT DO_STATUS FROM WMS_DELV_ORDER M " &
                    "WHERE DO_STATUS = 'POSTED' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "

        checkdt = gDB.getDataTable(SQLString)

        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The DO status is Posted, please re-open before post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is Posted, please re-open before post DO!", Session("gLang"))
            End If
            Return False
        End If

        For Each rows As DataRow In dt.Rows
            If rows.Item("mFlag") = "D" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Please save the DO first before Post", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", "Please save the DO first before Post", Session("gLang"))
                End If
                Return False
            End If
        Next

        If GridView1.Rows.Count <= 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "没有可供发布的物件!", Session("gLang"))
            End If
            Return False
        End If

        If Not save("Y") Then
            Return False
        End If

        If Not isValidPost() Then
            Return False
        End If

        Return True
    End Function

    'Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
    '    Dim updateSql, selectSql, SQLString As String
    '    Dim pl_dt As DataTable
    '    Dim checkdt As DataTable
    '    Dim itmKey As String
    '    Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
    '    Dim gConn As SqlConnection
    '    Dim transaction As SqlTransaction
    '    Dim lCOD_SEQ As String
    '    Dim successFlag As Boolean

    '    SQLString = "SELECT CO_STATUS FROM WMS_CUST_ORDER M " & _
    '        "WHERE CO_STATUS = 'CLOSED' " & _
    '        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' "
    '    checkdt = gDB.getDataTable(SQLString)
    '    If checkdt.Rows.Count > 0 Then
    '        If Session("gLang") = "E" Then
    '            uiFun.displayMsgNew(updtPnlAlert, "", "The SIR status is Closed, please re-open before post DO!", Session("gLang"))
    '        Else
    '            uiFun.displayMsgNew(updtPnlAlert, "", "The SIR status is Closed, please re-open before post DO!", Session("gLang"))
    '        End If
    '        Exit Sub
    '    End If

    '    SQLString = "SELECT DO_STATUS FROM WMS_DELV_ORDER M " & _
    '                "WHERE DO_STATUS = 'POSTED' " & _
    '                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
    '    checkdt = gDB.getDataTable(SQLString)
    '    If checkdt.Rows.Count > 0 Then
    '        If Session("gLang") = "E" Then
    '            uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is Posted, please re-open before post DO!", Session("gLang"))
    '        Else
    '            uiFun.displayMsgNew(updtPnlAlert, "", "The WIT status is Posted, please re-open before post DO!", Session("gLang"))
    '        End If
    '        Exit Sub
    '    End If

    '    For Each rows As DataRow In dt.Rows
    '        If rows.Item("mFlag") = "D" Then
    '            If Session("gLang") = "E" Then
    '                uiFun.displayMsgNew(updtPnlAlert, "", "Please save the WIT first before Post", Session("gLang"))
    '            Else
    '                uiFun.displayMsgNew(updtPnlAlert, "", "Please save the WIT first before Post", Session("gLang"))
    '            End If
    '            Exit Sub
    '        End If
    '    Next

    '    If GridView1.Rows.Count > 0 Then
    '        'If validateAll("Y") Then
    '        If save("Y") Then
    '            If isValidPost() Then
    '                gConn = gDB.getConnection()

    '                transaction = gConn.BeginTransaction()

    '                Try
    '                    qtyDict = New Dictionary(Of String, Double)
    '                    cbmDict = New Dictionary(Of String, Double)
    '                    wgtDict = New Dictionary(Of String, Double)

    '                    For Each rows As DataRow In dt.Rows
    '                        itmKey = gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" & _
    '                                gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" & _
    '                                gU.decodeNull(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" & _
    '                                gU.decodeNull(rows.Item("dod_batch_no").ToString.Trim, "")

    '                        If qtyDict.ContainsKey(itmKey) Then
    '                            qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0)
    '                            cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0), 14)
    '                            wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0), 14)
    '                        Else
    '                            qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0))
    '                            cbmDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0))
    '                            wgtDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0))
    '                        End If
    '                    Next

    '                    pl_dt = Session("_M_OB_DO_TMP_pl_dt")

    '                    Dim MANU_DATE As String = ""
    '                    Dim EXP_DATE As String = ""

    '                    'Sum total foi qty and picked qty for check stock bal validation
    '                    Dim objPlt As PickListTable
    '                    objPlt = New PickListTable(DO_CO_CODE.Text, DO_CODE.Text, STORER_CODE.SelectedValue, Session("IMP_CODE"))
    '                    objPlt.updateTotalQty(pl_dt)
    '                    objPlt.updateHoldBal(pl_dt)
    '                    objPlt = Nothing

    '                    'Check total bal (for hold stock, becoz hold stock hv no loc)
    '                    For Each rows As DataRow In pl_dt.Rows
    '                        If CDbl(rows.Item("hold_qty")) > 0 Then
    '                            If CDbl(rows.Item("total_item_qty")) > CDbl(rows.Item("avail_qty")) Then
    '                                If Session("gLang") = "E" Then
    '                                    uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                      "don\'t have enough total stock to delivery!", Session("gLang"))
    '                                Else
    '                                    uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                      "沒有足夠總貨存出貨!", Session("gLang"))
    '                                End If
    '                                transaction.Rollback()
    '                                gConn.Close()
    '                                gConn.Dispose()
    '                                Exit Sub
    '                            End If
    '                        End If
    '                    Next

    '                    Dim haveReDrumYN As String = "N"
    '                    'Check have redrum
    '                    For Each rows As DataRow In pl_dt.Rows
    '                        If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
    '                            haveReDrumYN = "Y"
    '                        End If
    '                    Next

    '                    If haveReDrumYN = "N" Then
    '                        uiFun.displayMsgNew(updtPnlAlert, "", "Re-Drum should be performed first!", Session("gLang"))
    '                        transaction.Rollback()
    '                        gConn.Close()
    '                        gConn.Dispose()
    '                        Exit Sub
    '                    End If

    '                    'Check for each loc bal
    '                    For Each rows As DataRow In pl_dt.Rows
    '                        Dim SrchStr As String = ""
    '                        SrchStr += "select ILOC_BAL_QTY, Convert(varchar, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
    '                        SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
    '                        SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
    '                        SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
    '                        SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
    '                        SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, ""), "000")) & "' "

    '                        If rows.Item("pld_batch_no").ToString.Trim <> "" Then
    '                            SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' "
    '                        Else
    '                            SrchStr += "AND ISNULL(ILOC_BATCH_NO,'') = '' "
    '                        End If

    '                        Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)


    '                        MANU_DATE = rows.Item("PLD_MANU_DATE").ToString.Trim
    '                        EXP_DATE = rows.Item("PLD_EXPIRY_DATE").ToString.Trim

    '                        If qtydt.Rows.Count > 0 Then
    '                            Dim stQTY As Integer = gU.decodeEmptyCInt(qtydt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
    '                            Dim plQTY As Integer = gU.decodeEmptyCInt(rows.Item("pld_item_qty").ToString.Trim, 0)

    '                            If stQTY < plQTY Then
    '                                If Session("gLang") = "E" Then
    '                                    uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                      "don\'t have enough stock to delivery!", Session("gLang"))
    '                                Else
    '                                    uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                      gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                      "沒有足夠貨存出貨!", Session("gLang"))
    '                                End If

    '                                transaction.Rollback()
    '                                gConn.Close()
    '                                gConn.Dispose()
    '                                Exit Sub
    '                            End If

    '                            If MANU_DATE = "" Then
    '                                MANU_DATE = qtydt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim
    '                            End If
    '                            If EXP_DATE = "" Then
    '                                EXP_DATE = qtydt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
    '                            End If
    '                        Else
    '                            If Session("gLang") = "E" Then
    '                                uiFun.displayMsgNew(updtPnlAlert, "", "Item " & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                  "don\'t have stock to delivery!", Session("gLang"))
    '                            Else
    '                                uiFun.displayMsgNew(updtPnlAlert, "", "物料" & gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & " | " & _
    '                                                  gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "") & " " & _
    '                                                  "沒有此貨存出貨!", Session("gLang"))
    '                            End If

    '                            transaction.Rollback()
    '                            gConn.Close()
    '                            gConn.Dispose()
    '                            Exit Sub
    '                        End If

    '                        st.IO_SEQ = ""
    '                        st.STORER_CODE = STORER_CODE.SelectedValue
    '                        st.ITM_CODE = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")
    '                        st.PACK_KEY = gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")
    '                        st.PALLET_NO = gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "000")
    '                        st.IO_CUST_CODE = CUS_CODE.SelectedValue
    '                        st.IO_WH = gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")
    '                        st.IO_AREA = gU.decodeNull(rows.Item("pld_area").ToString.Trim, "")
    '                        st.IO_LOC = gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")
    '                        st.IO_DOC = "DO"
    '                        st.IO_DOC_ID = DO_CODE.Text.Trim
    '                        st.IO_QTY = gU.decodeNull(rows.Item("pld_item_qty").ToString.Trim, "")
    '                        st.lO_BATCH_NO = gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")
    '                        st.IO_MANU_DATE = MANU_DATE
    '                        st.IO_EXPIRY_DATE = EXP_DATE


    '                        'Dim tempUOM2, tempQty2 As String

    '                        'tempUOM2 = ""
    '                        'tempQty2 = ""

    '                        'For Each d_rows As DataRow In dt.Rows
    '                        '    If d_rows.Item("dod_itm_code").ToString.Trim = rows.Item("pld_item_no").ToString.Trim AndAlso d_rows.Item("dod_pack_key").ToString.Trim = rows.Item("pld_pack_key").ToString.Trim AndAlso _
    '                        '              d_rows.Item("dod_pallet_no").ToString.Trim = rows.Item("pld_pallet_no").ToString.Trim AndAlso d_rows.Item("dod_batch_no").ToString.Trim = rows.Item("pld_batch_no").ToString.Trim Then

    '                        '        tempUOM2 = d_rows.Item("DOD_UOM2").ToString.Trim
    '                        '        'tempQty2 = d_rows.Item("DOD_QTY2").ToString.Trim

    '                        '        Exit For
    '                        '    End If
    '                        'Next

    '                        st.IOS_QTY2 = rows.Item("pld_qty2").ToString.Trim

    '                        st.IOS_SERIAL_NO = rows.Item("pld_serial_no").ToString.Trim

    '                        st.setOrgSerialInfo(rows.Item("pld_serial_no").ToString.Trim, gConn, transaction)

    '                        'st.IOS_DRUM_LEVEL = "1"
    '                        'st.IOS_UOM2 = tempUOM2
    '                        'st.IOS_ORG_QTY2 = ""
    '                        'st.IOS_QTY2 = ""
    '                        'st.IOS_DRUM_ID = ""
    '                        'st.IOS_ORG_SERIAL_NO = ""
    '                        'st.IOS_SERIAL_NO = ""


    '                        itmKey = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" & _
    '                                 gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" & _
    '                                 gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" & _
    '                                 gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")

    '                        If qtyDict.ContainsKey(itmKey) Then
    '                            Dim tmpCbm As Decimal = 0
    '                            Dim tmpwgt As Decimal = 0
    '                            Dim tmpqty As Decimal = 0

    '                            tmpCbm = cbmDict.Item(itmKey)
    '                            tmpwgt = wgtDict.Item(itmKey)
    '                            tmpqty = qtyDict.Item(itmKey)

    '                            If tmpqty = 0 Then tmpqty = 1

    '                            st.IO_CBM = Math.Round(tmpCbm / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
    '                            st.IO_KG = Math.Round(tmpwgt / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
    '                            'st.IO_CBM = Math.Round(tmpCbm * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
    '                            'st.IO_KG = Math.Round(tmpwgt * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
    '                        Else
    '                            st.IO_CBM = 0
    '                            st.IO_KG = 0
    '                        End If

    '                        If rows.Item("pld_serial_no").ToString.Trim = "" Then
    '                            updateSql = "update WMS_CUST_ORDER_D " & _
    '                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
    '                                            "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
    '                                            "sys_lud = Getdate() " & _
    '                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
    '                                        "AND ISNULL(COD_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', '') " & _
    '                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
    '                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

    '                            gDB.amendData(updateSql, gConn, transaction)

    '                        Else
    '                            selectSql = "select COD_SEQ as value " & _
    '                                        "from WMS_CUST_ORDER_D " & _
    '                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
    '                                        "AND ISNULL(COD_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', '') " & _
    '                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
    '                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " & _
    '                                        "AND ISNULL(COD_POST_QTY, 0) < ISNULL(COD_QTY, 0) " & _
    '                                        "order by convert(int, COD_SEQ) "

    '                            lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

    '                            updateSql = "update WMS_CUST_ORDER_D " & _
    '                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " & _
    '                                            "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " & _
    '                                            "sys_lud = Getdate() " & _
    '                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                        "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

    '                            gDB.amendData(updateSql, gConn, transaction)

    '                        End If

    '                        'Update coh_rel_qty for Hold stock logic
    '                        updateSql = "update wms_cust_order_hold " & _
    '                                    "set COH_REL_QTY = case when ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) < ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " then wms_cust_order_hold.coh_in_stock_qty else ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " & _
    '                                        "sys_lub = '" & Session("usr_id") & "', " & _
    '                                        "sys_lud = Getdate() " & _
    '                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                    "and wms_cust_order_hold.coh_status <> 'RELEASE' " & _
    '                                    "and ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) > 0 " & _
    '                                    "and ISNULL(wms_cust_order_hold.COH_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " & _
    '                                    "and ISNULL(wms_cust_order_hold.COH_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', '') " & _
    '                                    "AND wms_cust_order_hold.COH_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " & _
    '                                    "AND wms_cust_order_hold.COH_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "' " & _
    '                                    "and exists (" & _
    '                                        "select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
    '                                        "where wms_cust_order_hold.imp_code = d.imp_code " & _
    '                                        "and wms_cust_order_hold.storer_code = d.storer_code " & _
    '                                        "and wms_cust_order_hold.co_code = d.co_code " & _
    '                                        "and wms_cust_order_hold.cod_seq = d.cod_seq " & _
    '                                        "and c.imp_code = d.imp_code " & _
    '                                        "and c.storer_code = d.storer_code " & _
    '                                        "and c.co_code = d.co_code " & _
    '                                        "and c.CO_STATUS not in ('CLOSED', 'CANCELLED')) "

    '                        gDB.amendData(updateSql, gConn, transaction)

    '                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

    '                        st.UpdateStockTrans("OUT", gConn, transaction)

    '                        st.UpdateStockBalTrans("OUT", gConn, transaction)

    '                        If rows.Item("pld_serial_no").ToString.Trim <> "" Then
    '                            st.UpdateStockSerialTrans("OUT", gConn, transaction)

    '                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
    '                        End If

    '                    Next

    '                    For Each rows As DataRow In pl_dt.Rows
    '                        If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
    '                            st.cableReDrum(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, gConn, transaction)
    '                        End If
    '                    Next


    '                    updateSql = "update WMS_CUST_ORDER " & _
    '                                "set CO_STATUS = 'CLOSED', " & _
    '                                    "sys_lub = '" & Session("usr_id") & "', " & _
    '                                    "sys_lud = Getdate() " & _
    '                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                "AND NOT EXISTS (" & _
    '                                    "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
    '                                    "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
    '                                    "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
    '                                    "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
    '                                    "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " & _
    '                                    "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

    '                    gDB.amendData(updateSql, gConn, transaction)

    '                    updateSql = "update WMS_CUST_ORDER " & _
    '                                "set CO_STATUS = 'PARTIAL', " & _
    '                                "sys_lub = '" & Session("usr_id") & "', " & _
    '                                "sys_lud = Getdate() " & _
    '                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
    '                                "AND EXISTS (" & _
    '                                    "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
    '                                    "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
    '                                    "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
    '                                    "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
    '                                    "AND D.COD_POST_QTY > 0) " & _
    '                                "AND EXISTS (" & _
    '                                    "SELECT 1 FROM WMS_CUST_ORDER_D D " & _
    '                                    "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " & _
    '                                    "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " & _
    '                                    "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
    '                                    "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " & _
    '                                    "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

    '                    gDB.amendData(updateSql, gConn, transaction)

    '                    updateSql = "update WMS_DELV_ORDER " & _
    '                                "set DO_STATUS = 'POSTED', " & _
    '                                "DO_POSTED_DATE = Getdate(), " & _
    '                                "DO_POSTED_BY = '" & Session("usr_id") & "', " & _
    '                                "sys_lub = '" & Session("usr_id") & "', " & _
    '                                "sys_lud = Getdate() " & _
    '                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
    '                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
    '                                "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "

    '                    gDB.amendData(updateSql, gConn, transaction)

    '                    transaction.Commit()
    '                    gConn.Close()

    '                    'DO_STATUS.Text = "POSTED"
    '                    'ar.sec_viewMode = "Y"
    '                    'btnPost.Visible = False
    '                    'btnUnPost.Visible = True
    '                    'btnUnPick.Visible = False
    '                    'btnPick.Visible = False

    '                    'exceptionEditList.Add("saveConDate")
    '                    'exceptionEditList.Add("DO_CONF_DELTIME")
    '                    'exceptionEditList.Add("DO_CONF_DELDATE")
    '                    'exceptionEditList.Add("ImageButton3")

    '                    'ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

    '                    'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

    '                    'saveBtn1.Visible = False
    '                    'saveBtn2.Visible = False
    '                    'saveConDate.Visible = True

    '                    'uiFun.displayMsg(Me, "1007", "", Session("gLang"))

    '                    'Call BindGV()

    '                    'selectItemBtn.Attributes("onclick") = "ItemLookUp('" & STORER_CODE.SelectedValue & "');"
    '                    'cSBBtn.Attributes("onclick") = "checkSB('" & STORER_CODE.SelectedValue & "');"
    '                    'cSBBtn.Enabled = True

    '                    successFlag = True

    '                    If gConn IsNot Nothing Then
    '                        If gConn.State = ConnectionState.Open Then
    '                            gConn.Close()
    '                            gConn.Dispose()
    '                        End If
    '                    End If

    '                Catch ex As Exception
    '                    transaction.Rollback()
    '                    successFlag = False
    '                    gConn.Close()
    '                    gConn.Dispose()
    '                    Throw ex
    '                End Try

    '                If successFlag = True Then
    '                    reloadPage("Record has been Posted successfully!")

    '                    'Dim rmtPost As New RemotePost
    '                    'rmtPost.Url = "DOMain.aspx"
    '                    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
    '                    'rmtPost.Add("DO_CODE", DO_CODE.Text)
    '                    'rmtPost.alertMsg = "Record has been Posted successfully!"
    '                    'rmtPost.Post()
    '                Else

    '                End If
    '            End If
    '        End If
    '        'Else
    '        '    If Session("gLang") = "E" Then
    '        '        uiFun.displayMsgNew(updtPnlAlert, "", "Pick list location cannot be empty! Please check pick list.", Session("gLang"))
    '        '    Else
    '        '        uiFun.displayMsgNew(updtPnlAlert, "", "取貨位置不能空白! 請檢查取貨單", Session("gLang"))
    '        '    End If
    '        'End If
    '    Else
    '        If Session("gLang") = "E" Then
    '            uiFun.displayMsgNew(updtPnlAlert, "", "No item can be posted!", Session("gLang"))
    '        Else
    '            uiFun.displayMsgNew(updtPnlAlert, "", "没有可供发布的物件!", Session("gLang"))
    '        End If
    '    End If
    'End Sub

    Protected Sub addCOtoDO()
        Dim i, j As Integer
        Dim selectSql As String
        Dim hldDt As DataTable
        Dim SQLString As String
        Dim CO_dt As DataTable
        Dim COItemList As String = ""
        Dim COListarray As String()
        Dim COSeqListarray As String()
        Dim allowGo As Boolean
        Dim qtyListArray As String()
        Dim qtyitemDict As Dictionary(Of String, String)
        Dim tmpQty As Double

        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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

            COItemList = ""

            allowGo = True
            COListarray = Split(COList.Value, ", ")
            COSeqListarray = Split(COSeqList.Value, ", ")
            qtyListArray = Split(qtyList.Value, ", ")

            qtyitemDict = New Dictionary(Of String, String)

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
                SQLString = "SELECT d.CO_CODE, d.COD_SEQ, d.COD_QTY, m.CO_REM, m.CO_SHIP_MODE, Convert(datetime,m.CO_TARGET_DELDATE," & DDFORMAT & ") as CO_TARGET_DELDATE, " &
                                "d.COD_DISP_SEQ, d.COD_PALLET_NO, d.COD_ITM_CODE, d.COD_PACK_KEY, d.COD_ITM_DESC, d.COD_PACKING, " &
                                "m.CO_TRACK_NO, d.COD_UOM, d.COD_PCS_UOM, d.COD_TOT_WGT, d.COD_TOT_CBM, d.COD_REM, d.COD_VND_CODE, d.COD_BATCH_NO, " &
                                "m.CUS_CODE, m.CUS_NAME, m.CO_ADDR1, m.CO_ADDR2, m.CO_ADDR3, m.CO_AREA_DEL, m.CO_REGION_DEL, m.CO_PROVINCE, m.CO_CITY, " &
                                "m.CO_COUNTRY_DEL, m.CO_CUS_CONT, m.CO_CUS_CONT_TEL, m.CO_TRACK_NO, m.CO_INV_NO, m.CO_CUS_REF_NO, i.ITM_SKU_NO, i.ITM_DESC, " &
                                "v.AITM_QTY_PER_CTN, v.AITM_VOL, v.CARTON_CBM, d.COD_TICKET_NO,m.CO_PROVINCE, m.CO_CITY, " &
                                "m.CO_SENDER, m.CO_SENDER_COUNTRY, m.CO_SENDER_PROVINCE, m.CO_SENDER_REGION, m.CO_SENDER_ADDR, m.CO_SENDER_TEL, " &
                                "d.COD_UOM2, d.COD_QTY2, d.COD_WH_CODE, " &
                                "isnull(hi.HOLD_QTY, 0) as HOLD_QTY, " &
                                "i.ITM_TYPE, i.ITM_SERIAL_NO_YN, " &
                                "Convert(datetime,d.COD_EXPIRY_DATE, " & DDFORMAT & ") as COD_EXPIRY_DATE, " &
                                "Convert(datetime,d.COD_MANU_DATE, " & DDFORMAT & ") as COD_MANU_DATE, m.CO_EDI_SIR_NO " &
                             "from WMS_CUST_ORDER m " &
                             "inner join WMS_CUST_ORDER_D d " &
                                 "ON d.CO_CODE = m.CO_CODE " &
                                 "and d.STORER_CODE = m.STORER_CODE " &
                                 "and d.IMP_CODE = m.IMP_CODE " &
                             "left outer join WMS_ITEM i " &
                                 "on d.IMP_CODE = i.IMP_CODE " &
                                 "and d.STORER_CODE = i.STORER_CODE " &
                                 "and d.COD_ITM_CODE = i.ITM_CODE " &
                                 "and d.COD_PACK_KEY = i.PACK_KEY " &
                             "left outer join V_ALT_VEND_ITEM v " &
                                 "on d.IMP_CODE = v.IMP_CODE " &
                                 "and d.STORER_CODE = v.STORER_CODE " &
                                 "and d.COD_ITM_CODE = v.ITM_CODE " &
                                 "and d.COD_PACK_KEY = v.PACK_KEY " &
                             "left outer join (" &
                                     "select IMP_CODE, STORER_CODE, CO_CODE, COD_SEQ, sum(COH_QTY) as HOLD_QTY " &
                                     "from WMS_CUST_ORDER_HOLD " &
                                     "group by IMP_CODE, STORER_CODE, CO_CODE, COD_SEQ) hi " &
                                 "on d.IMP_CODE = hi.IMP_CODE " &
                                 "and d.STORER_CODE = hi.STORER_CODE " &
                                 "and d.CO_CODE = hi.CO_CODE " &
                                 "and d.COD_SEQ = hi.COD_SEQ " &
                             "where d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                             "and d.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                             "and d.CO_CODE + '_000_' + d.COD_SEQ IN ('" & Replace(COItemList, ", ", "', '") & "') "

                SQLString = SQLString & " order by d.COD_DISP_SEQ"
                REM **********************
                CO_dt = gDB.getDataTable(SQLString)

                If CO_dt.Rows.Count = 0 Then
                    Exit Sub
                Else
                    'CUS_CODE.SelectedValue = CO_dt.Rows(0).Item("CUS_CODE").ToString
                    CUS_NAME.Text = CO_dt.Rows(0).Item("CUS_NAME").ToString
                    DO_ADDR1.Text = CO_dt.Rows(0).Item("CO_ADDR1").ToString
                    DO_ADDR2.Text = CO_dt.Rows(0).Item("CO_ADDR2").ToString
                    DO_ADDR3.Text = CO_dt.Rows(0).Item("CO_ADDR3").ToString
                    DO_AREA_DEL.Text = CO_dt.Rows(0).Item("CO_AREA_DEL").ToString
                    DO_REGION_DEL.Text = CO_dt.Rows(0).Item("CO_REGION_DEL").ToString
                    DO_COUNTRY_DEL.Text = CO_dt.Rows(0).Item("CO_COUNTRY_DEL").ToString
                    DO_CUS_CONT.Text = CO_dt.Rows(0).Item("CO_CUS_CONT").ToString
                    DO_CUS_CONT_TEL.Text = CO_dt.Rows(0).Item("CO_CUS_CONT_TEL").ToString
                    'DO_TRACK_NO.Text = CO_dt.Rows(0).Item("CO_TRACK_NO").ToString
                    DO_SHIP_MODE.SelectedValue = CO_dt.Rows(0).Item("CO_SHIP_MODE").ToString
                    DO_TARGET_DELDATE.Text = CO_dt.Rows(0).Item("CO_TARGET_DELDATE").ToString
                    DO_INV_NO.Text = CO_dt.Rows(0).Item("CO_INV_NO").ToString
                    DO_REM.Text = CO_dt.Rows(0).Item("CO_REM").ToString
                    DO_CUS_REF_NO.Text = CO_dt.Rows(0).Item("CO_CUS_REF_NO").ToString
                    DO_EDI_SIR_NO.Text = CO_dt.Rows(0).Item("CO_EDI_SIR_NO").ToString

                    DO_PROVINCE.Text = CO_dt.Rows(0).Item("CO_PROVINCE").ToString
                    DO_CITY.Text = CO_dt.Rows(0).Item("CO_CITY").ToString

                    'm.CO_SENDER, m.CO_SENDER_COUNTRY, m.CO_SENDER_PROVINCE, m.CO_SENDER_REGION, m.CO_SENDER_ADDR, m.CO_SENDER_TEL,
                    DO_SENDER.Text = CO_dt.Rows(0).Item("CO_SENDER").ToString
                    DO_SENDER_COUNTRY.Text = CO_dt.Rows(0).Item("CO_SENDER_COUNTRY").ToString
                    DO_SENDER_PROVINCE.Text = CO_dt.Rows(0).Item("CO_SENDER_PROVINCE").ToString
                    DO_SENDER_REGION.Text = CO_dt.Rows(0).Item("CO_SENDER_REGION").ToString
                    DO_SENDER_ADDR.Text = CO_dt.Rows(0).Item("CO_SENDER_ADDR").ToString
                    DO_SENDER_TEL.Text = CO_dt.Rows(0).Item("CO_SENDER_TEL").ToString

                    DO_CONSIGNEE.Text = CO_dt.Rows(0).Item("CO_CUS_CONT").ToString
                    DO_CONSIGNEE_ADDR1.Text = CO_dt.Rows(0).Item("CO_ADDR1").ToString

                    'CO_SENDER, CO_SENDER_COUNTRY, CO_SENDER_PROVINCE, CO_SENDER_REGION, CO_SENDER_ADDR, CO_SENDER_TEL

                End If

                For i = 0 To CO_dt.Rows.Count - 1

                    Dim itmQty As Integer = 0

                    If qtyitemDict.ContainsKey(CO_dt.Rows(i).Item("CO_CODE").ToString & "_000_" & CO_dt.Rows(i).Item("COD_SEQ").ToString) Then
                        itmQty = gU.decodeEmptyCInt(qtyitemDict(CO_dt.Rows(i).Item("CO_CODE").ToString & "_000_" & CO_dt.Rows(i).Item("COD_SEQ").ToString), CInt(CO_dt.Rows(i).Item("COD_QTY")))
                    Else
                        Dim qtySQL As String = ""
                        Dim qtyDT As DataTable

                        qtySQL = "SELECT WMS_CUST_ORDER.CO_CODE, WMS_CUST_ORDER_D.COD_SEQ, " &
                                     "WMS_CUST_ORDER_D.COD_QTY,  " &
                                     "Case sign((ISNULL(WMS_CUST_ORDER_D.COD_QTY, 0) - ISNULL(delv.dod_qty,0))) when -1 then 0 else (ISNULL(WMS_CUST_ORDER_D.COD_QTY,0) - ISNULL(delv.dod_qty,0)) end as avai_qty " &
                                 "from WMS_CUST_ORDER " &
                                 "inner join WMS_CUST_ORDER_D " &
                                     "on WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE " &
                                     "and WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE " &
                                     "and WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE " &
                                 "left outer join ( " &
                                         "select d.IMP_CODE, d.STORER_CODE, h.DO_CO_CODE, d.DOD_ITM_CODE, " &
                                         "d.DOD_PACK_KEY, d.DOD_PALLET_NO, d.DOD_BATCH_NO, sum(d.dod_qty) as dod_qty " &
                                         "from WMS_DELV_ORDER h, WMS_DELV_ORDER_D d " &
                                         "where h.IMP_CODE = d.IMP_CODE " &
                                         "and h.STORER_CODE = d.STORER_CODE " &
                                         "and h.DO_CODE = d.DO_CODE " &
                                         "and ISNULL(h.DO_STATUS, '') <> 'CANCELLED' " &
                                         "group by d.IMP_CODE, d.STORER_CODE, h.DO_CO_CODE, d.DOD_ITM_CODE, d.DOD_PACK_KEY, d.DOD_PALLET_NO, d.DOD_BATCH_NO) delv " &
                                     "on WMS_CUST_ORDER_D.IMP_CODE = delv.IMP_CODE  " &
                                     "and WMS_CUST_ORDER_D.STORER_CODE = delv.STORER_CODE  " &
                                     "and WMS_CUST_ORDER_D.CO_CODE = delv.DO_CO_CODE  " &
                                     "and WMS_CUST_ORDER_D.COD_ITM_CODE = delv.DOD_ITM_CODE  " &
                                     "and WMS_CUST_ORDER_D.COD_PACK_KEY = delv.DOD_PACK_KEY  " &
                                     "and ISNULL(WMS_CUST_ORDER_D.COD_PALLET_NO, '000') = ISNULL(delv.DOD_PALLET_NO , '000') " &
                                     "and ISNULL(WMS_CUST_ORDER_D.COD_BATCH_NO, '') = ISNULL(delv.DOD_BATCH_NO , '') " &
                                 "WHERE WMS_CUST_ORDER.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                 "and WMS_CUST_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                                 "and WMS_CUST_ORDER.CO_CODE='" & gU.dbEncode(CO_dt.Rows(i).Item("CO_CODE").ToString) & "' " &
                                 "AND WMS_CUST_ORDER_D.COD_SEQ='" & gU.dbEncode(CO_dt.Rows(i).Item("COD_SEQ").ToString) & "' " &
                                 "and WMS_CUST_ORDER.CO_STATUS <> 'CLOSED' "


                        qtyDT = gDB.getDataTable(qtySQL)

                        If qtyDT.Rows.Count > 0 Then
                            itmQty = qtyDT.Rows(0).Item("avai_qty").ToString.Trim
                        Else
                            itmQty = CInt(CO_dt.Rows(i).Item("COD_QTY"))
                        End If

                    End If

                    If CO_dt.Rows(i).Item("COD_BATCH_NO").ToString.Trim <> "" OrElse CO_dt.Rows(i).Item("HOLD_QTY").ToString.Trim = "0" Then
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
                        dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = CO_dt.Rows(i).Item("COD_DISP_SEQ")
                        dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = CO_dt.Rows(i).Item("COD_PALLET_NO")
                        'dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = gU.decodeNullOrEmpty(CO_dt.Rows(i).Item("COD_PALLET_NO"), "000")
                        dt.Rows(rows_count - 1).Item("DOD_CARTON_NO") = ""
                        dt.Rows(rows_count - 1).Item("DOD_ITM_CODE") = CO_dt.Rows(i).Item("COD_ITM_CODE")
                        dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CO_dt.Rows(i).Item("ITM_SKU_NO")

                        dt.Rows(rows_count - 1).Item("DOD_PACK_KEY") = CO_dt.Rows(i).Item("COD_PACK_KEY")
                        dt.Rows(rows_count - 1).Item("DOD_ITM_DESC") = CO_dt.Rows(i).Item("COD_ITM_DESC")
                        dt.Rows(rows_count - 1).Item("ITM_DESC") = CO_dt.Rows(i).Item("ITM_DESC")
                        dt.Rows(rows_count - 1).Item("DOD_PACK_NO") = ""
                        dt.Rows(rows_count - 1).Item("DOD_PACK_TYPE") = CO_dt.Rows(i).Item("COD_PACKING")
                        dt.Rows(rows_count - 1).Item("DOD_TICKET_NO") = CO_dt.Rows(i).Item("COD_TICKET_NO")
                        dt.Rows(rows_count - 1).Item("DOD_WH_CODE") = CO_dt.Rows(i).Item("COD_WH_CODE")


                        'dt.Rows(rows_count - 1).Item("DOD_QTY") = CO_dt.Rows(i).Item("COD_QTY")
                        dt.Rows(rows_count - 1).Item("DOD_QTY") = itmQty

                        dt.Rows(rows_count - 1).Item("DOD_UOM") = CO_dt.Rows(i).Item("COD_UOM")
                        dt.Rows(rows_count - 1).Item("DOD_PCS_UOM") = CO_dt.Rows(i).Item("COD_PCS_UOM")

                        'dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = itmQty * CInt(CO_dt.Rows(i).Item("COD_PCS_UOM"))
                        dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = itmQty * If(CO_dt.Rows(i).Item("COD_PCS_UOM") IsNot DBNull.Value, CO_dt.Rows(i).Item("COD_PCS_UOM").ToString(), "0")

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


                        dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = CO_dt.Rows(i).Item("ITM_SERIAL_NO_YN")
                        dt.Rows(rows_count - 1).Item("ITM_TYPE") = CO_dt.Rows(i).Item("ITM_TYPE")

                        'dt.Rows(rows_count - 1).Item("DOD_NO_OF_CARTON") = 1


                        'CUS_CODE.SelectedValue = CO_dt.Rows(i).Item("CUS_CODE").ToString
                        'CUS_NAME.Text = CO_dt.Rows(i).Item("CUS_NAME").ToString
                        'DO_ADDR1.Text = CO_dt.Rows(i).Item("CO_ADDR1").ToString
                        'DO_ADDR2.Text = CO_dt.Rows(i).Item("CO_ADDR2").ToString
                        'DO_ADDR3.Text = CO_dt.Rows(i).Item("CO_ADDR3").ToString
                        'DO_AREA_DEL.Text = CO_dt.Rows(i).Item("CO_AREA_DEL").ToString
                        'DO_REGION_DEL.Text = CO_dt.Rows(i).Item("CO_REGION_DEL").ToString
                        'DO_COUNTRY_DEL.Text = CO_dt.Rows(i).Item("CO_COUNTRY_DEL").ToString
                        'DO_CUS_CONT.Text = CO_dt.Rows(i).Item("CO_CUS_CONT").ToString
                        'DO_CUS_CONT_TEL.Text = CO_dt.Rows(i).Item("CO_CUS_CONT_TEL").ToString
                        'DO_TRACK_NO.Text = CO_dt.Rows(i).Item("CO_TRACK_NO").ToString
                        'DO_SHIP_MODE.SelectedValue = CO_dt.Rows(i).Item("CO_SHIP_MODE").ToString
                        'DO_TARGET_DELDATE.Text = CO_dt.Rows(i).Item("CO_TARGET_DELDATE").ToString
                        'DO_INV_NO.Text = CO_dt.Rows(i).Item("CO_INV_NO").ToString
                        'DO_REM.Text = CO_dt.Rows(i).Item("CO_REM").ToString
                        'DO_CUS_REF_NO.Text = CO_dt.Rows(i).Item("CO_CUS_REF_NO").ToString
                        REM **********************
                        dt.Rows(rows_count - 1).Item("mFlag") = "N"
                    Else


                        selectSql = "select COH_BATCH_NO, COH_PALLET_NO, isnull(COH_IN_STOCK_QTY, 0) - isnull(COH_REL_QTY, 0) as hold_qty " &
                                    "from WMS_CUST_ORDER_HOLD " &
                                    "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and IMP_CODE = '" & Session("IMP_CODE") & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(CO_dt.Rows(i).Item("CO_CODE").ToString) & "' " &
                                    "and COD_SEQ = '" & gU.dbEncode(CO_dt.Rows(i).Item("COD_SEQ").ToString) & "' " &
                                    "order by COH_SEQ "

                        hldDt = gDB.getDataTable(selectSql)

                        'Loop hold records to generate detail with batch no.
                        For j = 0 To hldDt.Rows.Count - 1
                            If CDbl(hldDt.Rows(j).Item("hold_qty").ToString.Trim) > 0 Then
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
                                dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = CO_dt.Rows(i).Item("COD_DISP_SEQ")
                                dt.Rows(rows_count - 1).Item("DOD_CARTON_NO") = ""
                                dt.Rows(rows_count - 1).Item("DOD_ITM_CODE") = CO_dt.Rows(i).Item("COD_ITM_CODE")
                                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CO_dt.Rows(i).Item("ITM_SKU_NO")
                                dt.Rows(rows_count - 1).Item("DOD_PACK_KEY") = CO_dt.Rows(i).Item("COD_PACK_KEY")
                                dt.Rows(rows_count - 1).Item("DOD_ITM_DESC") = CO_dt.Rows(i).Item("COD_ITM_DESC")
                                dt.Rows(rows_count - 1).Item("ITM_DESC") = CO_dt.Rows(i).Item("ITM_DESC")
                                dt.Rows(rows_count - 1).Item("DOD_PACK_NO") = ""
                                dt.Rows(rows_count - 1).Item("DOD_PACK_TYPE") = CO_dt.Rows(i).Item("COD_PACKING")
                                dt.Rows(rows_count - 1).Item("DOD_TICKET_NO") = CO_dt.Rows(i).Item("COD_TICKET_NO")
                                dt.Rows(rows_count - 1).Item("DOD_WH_CODE") = CO_dt.Rows(i).Item("COD_WH_CODE")

                                If itmQty >= CDbl(hldDt.Rows(j).Item("hold_qty").ToString.Trim) Then
                                    tmpQty = CDbl(hldDt.Rows(j).Item("hold_qty").ToString.Trim)
                                    itmQty = itmQty - tmpQty
                                Else
                                    tmpQty = itmQty
                                    itmQty = 0
                                End If

                                dt.Rows(rows_count - 1).Item("DOD_QTY") = tmpQty

                                dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = tmpQty * CInt(CO_dt.Rows(i).Item("COD_PCS_UOM"))

                                If tmpQty <> 0 Then
                                    dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("AITM_VOL"), "0"), 0) * tmpQty
                                    dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = gU.decodeEmptyCdbl(DB.decodeDBNull(CO_dt.Rows(i).Item("CARTON_CBM"), "0"), 0) * tmpQty
                                Else
                                    dt.Rows(rows_count - 1).Item("DOD_TOT_WGT") = CO_dt.Rows(i).Item("AITM_VOL")
                                    dt.Rows(rows_count - 1).Item("DOD_TOT_CBM") = CO_dt.Rows(i).Item("CARTON_CBM")
                                End If

                                dt.Rows(rows_count - 1).Item("DOD_UOM") = CO_dt.Rows(i).Item("COD_UOM")
                                dt.Rows(rows_count - 1).Item("DOD_PCS_UOM") = CO_dt.Rows(i).Item("COD_PCS_UOM")
                                dt.Rows(rows_count - 1).Item("DOD_QTY2") = CO_dt.Rows(i).Item("COD_QTY2")
                                dt.Rows(rows_count - 1).Item("DOD_UOM2") = CO_dt.Rows(i).Item("COD_UOM2")
                                dt.Rows(rows_count - 1).Item("DOD_REM") = CO_dt.Rows(i).Item("COD_REM")
                                dt.Rows(rows_count - 1).Item("DOD_VND_CODE") = CO_dt.Rows(i).Item("COD_VND_CODE")
                                dt.Rows(rows_count - 1).Item("DOD_EXPIRY_DATE") = CO_dt.Rows(i).Item("COD_EXPIRY_DATE")
                                dt.Rows(rows_count - 1).Item("DOD_MANU_DATE") = CO_dt.Rows(i).Item("COD_MANU_DATE")

                                dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = CO_dt.Rows(i).Item("ITM_SERIAL_NO_YN")
                                dt.Rows(rows_count - 1).Item("ITM_TYPE") = CO_dt.Rows(i).Item("ITM_TYPE")

                                dt.Rows(rows_count - 1).Item("DOD_BATCH_NO") = hldDt.Rows(j).Item("COH_BATCH_NO").ToString.Trim
                                dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = hldDt.Rows(j).Item("COH_PALLET_NO").ToString.Trim

                                REM **********************
                                dt.Rows(rows_count - 1).Item("mFlag") = "N"
                            End If

                            If itmQty <= 0 Then
                                Exit For
                            End If
                        Next

                        'If there is qty left (hold qty < itm qty)
                        If itmQty > 0 Then
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
                            dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = CO_dt.Rows(i).Item("COD_DISP_SEQ")
                            dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = CO_dt.Rows(i).Item("COD_PALLET_NO")
                            'dt.Rows(rows_count - 1).Item("DOD_PALLET_NO") = gU.decodeNullOrEmpty(CO_dt.Rows(i).Item("COD_PALLET_NO"), "000")
                            dt.Rows(rows_count - 1).Item("DOD_CARTON_NO") = ""
                            dt.Rows(rows_count - 1).Item("DOD_ITM_CODE") = CO_dt.Rows(i).Item("COD_ITM_CODE")
                            dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CO_dt.Rows(i).Item("ITM_SKU_NO")
                            dt.Rows(rows_count - 1).Item("DOD_PACK_KEY") = CO_dt.Rows(i).Item("COD_PACK_KEY")
                            dt.Rows(rows_count - 1).Item("DOD_ITM_DESC") = CO_dt.Rows(i).Item("COD_ITM_DESC")
                            dt.Rows(rows_count - 1).Item("ITM_DESC") = CO_dt.Rows(i).Item("ITM_DESC")
                            dt.Rows(rows_count - 1).Item("DOD_PACK_NO") = ""
                            dt.Rows(rows_count - 1).Item("DOD_PACK_TYPE") = CO_dt.Rows(i).Item("COD_PACKING")
                            dt.Rows(rows_count - 1).Item("DOD_TICKET_NO") = CO_dt.Rows(i).Item("COD_TICKET_NO")
                            dt.Rows(rows_count - 1).Item("DOD_WH_CODE") = CO_dt.Rows(i).Item("COD_WH_CODE")

                            dt.Rows(rows_count - 1).Item("DOD_QTY") = itmQty

                            dt.Rows(rows_count - 1).Item("DOD_UOM") = CO_dt.Rows(i).Item("COD_UOM")
                            dt.Rows(rows_count - 1).Item("DOD_PCS_UOM") = CO_dt.Rows(i).Item("COD_PCS_UOM")

                            dt.Rows(rows_count - 1).Item("DOD_TOTPCS") = itmQty * CInt(CO_dt.Rows(i).Item("COD_PCS_UOM"))

                            dt.Rows(rows_count - 1).Item("DOD_QTY2") = CO_dt.Rows(i).Item("COD_QTY2")
                            dt.Rows(rows_count - 1).Item("DOD_UOM2") = CO_dt.Rows(i).Item("COD_UOM2")

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

                            dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = CO_dt.Rows(i).Item("ITM_SERIAL_NO_YN")
                            dt.Rows(rows_count - 1).Item("ITM_TYPE") = CO_dt.Rows(i).Item("ITM_TYPE")

                            REM **********************
                            dt.Rows(rows_count - 1).Item("mFlag") = "N"
                        End If

                    End If


                Next
                dt.AcceptChanges()
                Session("dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()




                '==========================================================================================================================================================================
                'No pre-gen of picking and packing ========================================================================================================================================
                '==========================================================================================================================================================================

                'Dim pickListDt, pl_dt As DataTable
                'Dim objPlt As PickListTable
                'Dim plSeq As Integer

                'objPlt = New PickListTable(DO_CO_CODE.Text, DO_CODE.Text, STORER_CODE.SelectedValue, Session("IMP_CODE"))

                'pickListDt = objPlt.generatePickList(dt)

                'objPlt = Nothing


                ''Dim LOC_dt, pickListDt, pl_dt As DataTable
                ' ''Dim objPlt As PickListTable
                ''Dim plSeq As Integer

                ''SQLString = "SELECT l.ITM_CODE, l.PACK_KEY, l.ILOC_PALLET_NO, d.COD_QTY, l.ILOC_BAL_QTY, l.ILOC_LOC, " & _
                ''                "l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO " & _
                ''            "from (" & _
                ''                "select COD_ITM_CODE, COD_PACK_KEY, ISNULL(COD_PALLET_NO, '000') as COD_PALLET_NO, sum(COD_QTY) as COD_QTY " & _
                ''                "from WMS_CUST_ORDER_D " & _
                ''                "where CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " & _
                ''                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                ''                "and IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                ''                "and ISNULL(COD_QTY, 0) > 0 " & _
                ''                "group by COD_ITM_CODE, COD_PACK_KEY, COD_PALLET_NO " & _
                ''                 ") d inner join WMS_ITEM_LOC_BAL l on " & _
                ''            "d.COD_ITM_CODE = l.ITM_CODE " & _
                ''            "and d.COD_PACK_KEY = l.PACK_KEY " & _
                ''            "and d.COD_PALLET_NO = l.ILOC_PALLET_NO " & _
                ''            "left outer join WMS_DATE_CODE DC on " & _
                ''            "l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
                ''            "and l.IMP_CODE = DC.IMP_CODE " & _
                ''            "and l.STORER_CODE = DC.STORER_CODE " & _
                ''            "where l.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                ''            "and l.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                ''            "and ISNULL(l.ILOC_BAL_QTY, 0) > 0 " & _
                ''            "order by l.ITM_CODE, DC.DC_CONV_DATE, l.PACK_KEY, l.ILOC_BATCH_NO, l.ILOC_BAL_QTY desc "

                ''LOC_dt = gDB.getDataTable(SQLString)

                ' ''objPlt = New PickListTable(LOC_dt)

                ' ''pickListDt = objPlt.getPickList()

                ' ''objPlt = Nothing

                ''pickListDt = New DataTable
                ''pickListDt.Columns.Add("PLD_ITEM_NO", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_PACK_KEY", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_PALLET_NO", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_DO_QTY", Type.GetType("System.Double"))
                ''pickListDt.Columns.Add("PLD_ITEM_QTY", Type.GetType("System.Double"))
                ''pickListDt.Columns.Add("ILOC_BAL_QTY", Type.GetType("System.Double"))
                ''pickListDt.Columns.Add("PLD_WH", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_LOC", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_FLOOR", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_AREA", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_RACK", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_BIN", Type.GetType("System.String"))
                ''pickListDt.Columns.Add("PLD_BATCH_NO", Type.GetType("System.String"))

                ''Dim currQty As Double = 0

                ''For i = 0 To LOC_dt.Rows.Count - 1
                ''    If gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0) > 0 AndAlso _
                ''        LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim <> "" Then
                ''        Dim itmK As String = LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
                ''                                                 LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
                ''                                                 LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString

                ''        Dim doRow As DataRow() = dt.Select("DOD_ITM_CODE + '#_#' + DOD_PACK_KEY + '#_#' + DOD_PALLET_NO = '" & _
                ''                                                        LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
                ''                                                          LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
                ''                                                          LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "'")
                ''        Dim reqQty As Double = 0

                ''        If doRow.Count > 0 Then
                ''            reqQty = gU.decodeEmptyCdbl(doRow(0).Item("DOD_QTY"), 0)
                ''        End If

                ''        If currQty < reqQty Then
                ''            Dim balQty As Double = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)

                ''            Dim pNewRow = pickListDt.NewRow
                ''            pNewRow.Item("PLD_ITEM_NO") = LOC_dt.Rows(i).Item("ITM_CODE").ToString.Trim
                ''            pNewRow.Item("PLD_PACK_KEY") = LOC_dt.Rows(i).Item("PACK_KEY").ToString.Trim
                ''            pNewRow.Item("PLD_PALLET_NO") = LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim

                ''            pNewRow.Item("PLD_WH") = LOC_dt.Rows(i).Item("ILOC_WH").ToString.Trim
                ''            pNewRow.Item("PLD_LOC") = LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                ''            pNewRow.Item("PLD_FLOOR") = LOC_dt.Rows(i).Item("ILOC_FLOOR").ToString.Trim
                ''            pNewRow.Item("PLD_AREA") = LOC_dt.Rows(i).Item("ILOC_AREA").ToString.Trim
                ''            pNewRow.Item("PLD_RACK") = LOC_dt.Rows(i).Item("ILOC_RACK").ToString.Trim
                ''            pNewRow.Item("PLD_BIN") = LOC_dt.Rows(i).Item("ILOC_BIN").ToString.Trim
                ''            pNewRow.Item("PLD_BATCH_NO") = LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                ''            pNewRow.Item("ILOC_BAL_QTY") = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)
                ''            pNewRow.Item("PLD_DO_QTY") = reqQty

                ''            If currQty + balQty >= reqQty Then
                ''                pNewRow.Item("PLD_ITEM_QTY") = reqQty - currQty
                ''                currQty = reqQty
                ''            Else

                ''                pNewRow.Item("PLD_ITEM_QTY") = balQty

                ''                currQty += balQty

                ''            End If

                ''            pickListDt.Rows.Add(pNewRow)
                ''            pickListDt.AcceptChanges()
                ''        End If
                ''    End If
                ''Next

                'pl_dt = Session("_M_OB_DO_TMP_pl_dt")
                'plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                'For i = 0 To pickListDt.Rows.Count - 1
                '    pl_dt.Rows.Add()

                '    plSeq = plSeq + 1

                '    rows_count = pl_dt.Rows.Count

                '    pl_dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                '    pl_dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = pickListDt.Rows(i).Item("PLD_ITEM_NO")
                '    pl_dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = pickListDt.Rows(i).Item("ITM_SKU_NO")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = pickListDt.Rows(i).Item("PLD_PACK_KEY")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = gU.decodeNullOrEmpty(pickListDt.Rows(i).Item("PLD_PALLET_NO"), "000")

                '    pl_dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = pickListDt.Rows(i).Item("PLD_FOI_QTY")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                '    pl_dt.Rows(rows_count - 1).Item("PLD_QTY2") = pickListDt.Rows(i).Item("PLD_QTY2")

                '    pl_dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = pickListDt.Rows(i).Item("PLD_DO_QTY")
                '    pl_dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = pickListDt.Rows(i).Item("ILOC_BAL_QTY")

                '    pl_dt.Rows(rows_count - 1).Item("HOLD_QTY") = pickListDt.Rows(i).Item("HOLD_QTY")
                '    pl_dt.Rows(rows_count - 1).Item("TOTAL_BAL") = pickListDt.Rows(i).Item("TOTAL_BAL")


                '    pl_dt.Rows(rows_count - 1).Item("ILOC_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")

                '    pl_dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = pickListDt.Rows(i).Item("ILOC_MANU_DATE")

                '    pl_dt.Rows(rows_count - 1).Item("PLD_WH") = pickListDt.Rows(i).Item("PLD_WH")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_LOC") = pickListDt.Rows(i).Item("PLD_LOC")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_ORG_LOC") = pickListDt.Rows(i).Item("PLD_LOC")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_REMARK") = ""
                '    pl_dt.Rows(rows_count - 1).Item("PLD_FLOOR") = pickListDt.Rows(i).Item("PLD_FLOOR")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_AREA") = pickListDt.Rows(i).Item("PLD_AREA")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_RACK") = pickListDt.Rows(i).Item("PLD_RACK")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_BIN") = pickListDt.Rows(i).Item("PLD_BIN")
                '    pl_dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                '    pl_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                '    pl_dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = pickListDt.Rows(i).Item("PLD_BATCH_NO")

                '    pl_dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = pickListDt.Rows(i).Item("ITM_SERIAL_NO_YN")
                'Next

                'Dim isItemExist As Boolean

                ''Check if any item is not existed in the pick list
                'For i = 0 To dt.Rows.Count - 1
                '    If dt.Rows(i).Item("mFlag").ToString <> "D" Then
                '        isItemExist = False
                '        For j = 0 To pl_dt.Rows.Count - 1
                '            If pl_dt.Rows(j).Item("mFlag").ToString <> "D" Then
                '                If dt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim = pl_dt.Rows(j).Item("PLD_ITEM_NO").ToString.Trim AndAlso _
                '                    dt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim = pl_dt.Rows(j).Item("PLD_PACK_KEY").ToString.Trim AndAlso _
                '                    dt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim = pl_dt.Rows(j).Item("PLD_PALLET_NO").ToString.Trim AndAlso _
                '                    (dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = "" OrElse dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = pl_dt.Rows(j).Item("PLD_BATCH_NO").ToString.Trim) Then

                '                    isItemExist = True
                '                    Exit For
                '                End If
                '            End If
                '        Next

                '        If Not isItemExist Then
                '            pl_dt.Rows.Add()

                '            plSeq = plSeq + 1

                '            rows_count = pl_dt.Rows.Count

                '            pl_dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                '            pl_dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                '            pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = dt.Rows(i).Item("DOD_ITM_CODE")
                '            pl_dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = dt.Rows(i).Item("DOD_PACK_KEY")
                '            pl_dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = gU.decodeNullOrEmpty(dt.Rows(i).Item("DOD_PALLET_NO"), "000")
                '            pl_dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = 0
                '            pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                '            pl_dt.Rows(rows_count - 1).Item("PLD_QTY2") = dt.Rows(i).Item("DOD_QTY2")

                '            pl_dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = dt.Rows(i).Item("DOD_QTY")

                '            pl_dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = 0
                '            pl_dt.Rows(rows_count - 1).Item("HOLD_QTY") = 0
                '            pl_dt.Rows(rows_count - 1).Item("TOTAL_BAL") = 0

                '            pl_dt.Rows(rows_count - 1).Item("PLD_WH") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_LOC") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_ORG_LOC") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_REMARK") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_FLOOR") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_AREA") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_RACK") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_BIN") = ""
                '            pl_dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                '            pl_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                '            pl_dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = dt.Rows(i).Item("DOD_BATCH_NO")

                '            pl_dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = dt.Rows(i).Item("DOD_EXPIRY_DATE")
                '            pl_dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = dt.Rows(i).Item("DOD_MANU_DATE")

                '            pl_dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = dt.Rows(i).Item("ITM_SERIAL_NO_YN")

                '        End If
                '    End If
                'Next

                'Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)
                'delPLZeroRow(pl_dt)

                'Dim dtldt As DataTable = Session("dt")
                'Dim pi_dt As DataTable = Session("_M_OB_DO_TMP_pi_dt")
                'Dim nextSEQ As String = ""

                'If dtldt IsNot Nothing AndAlso dtldt.Rows.Count > 0 Then

                '    If pi_dt Is Nothing AndAlso pi_dt.Rows.Count = 0 Then
                '        pi_dt = New DataTable
                '        '" '"select pd.*, pd.pad_pack_key as pad_pack_size, pi.pld_item_qty, 'U' as mflag " & _
                '        Dim SQLStr As String = "select pd.*, pi.pld_foi_qty, pi.pld_item_qty, 'U' as mflag " & _
                '       " from wms_do_packing_d pd inner join WMS_DO_PICKLIST_D pi on " & _
                '       "pd.imp_code = pi.imp_code " & _
                '       "and pd.storer_code = pi.storer_code " & _
                '       "and pd.pad_pack_key= pi.pld_pack_key " & _
                '       "and pd.pad_itm_code = pi.pld_item_no " & _
                '       "and pd.do_code = pi.do_code " & _
                '        "and 1 = 0"

                '        gDB.getDataTable(SQLStr, , , pi_dt)
                '    Else
                '        For Each piRow As DataRow In pi_dt.Rows
                '            piRow.Item("mFlag") = "D"
                '        Next

                '        pi_dt.AcceptChanges()
                '    End If

                '    If nextSEQ = "" Then
                '        Dim seqString As String = "select MAX(CAST(pad_pack_no AS int)) + 1 from wms_do_packing_d " & _
                '                    "where IMP_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("IMP_CODE"))) & "' " & _
                '                    "and STORER_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code").ToString)) & "' " & _
                '                    "and DO_CODE = '" & gU.dbEncode(dtldt.Rows(0).Item("DO_CODE").ToString) & "' "

                '        nextSEQ = DB.getValueFromSQL(seqString)

                '        If nextSEQ Is Nothing OrElse nextSEQ = "" Then
                '            nextSEQ = "1"
                '        End If
                '    End If

                '    Dim indx As Integer = 1

                '    For Each row As DataRow In pickListDt.Rows
                '        Dim isPIItemExist As Boolean = False

                '        If Not isPIItemExist Then
                '            Dim newRow = pi_dt.NewRow

                '            newRow.Item("pad_pack_no") = nextSEQ
                '            newRow.Item("PAD_DISPLAY_SEQ") = indx

                '            newRow.Item("pad_pallet_no") = gU.decodeNullOrEmpty(row.Item("pld_pallet_no"), "000")
                '            newRow.Item("pad_itm_code") = row.Item("PLD_ITEM_NO")
                '            newRow.Item("pad_batch_no") = row.Item("pld_batch_no")
                '            newRow.Item("pad_pack_key") = row.Item("pld_pack_key")
                '            'newRow.Item("pad_qty") = row.Item("pld_item_qty")
                '            newRow.Item("pad_qty") = row.Item("pld_foi_qty")
                '            'newRow.Item("pad_qty") = 0
                '            newRow.Item("pad_pack_by") = Session("usr_id").ToString

                '            newRow.Item("imp_code") = Server.UrlDecode(Request("IMP_CODE"))
                '            newRow.Item("storer_code") = Server.UrlDecode(Request("storer_code"))

                '            newRow.Item("mFlag") = "N"

                '            Dim dtRows As DataRow() = dtldt.Select("dod_itm_code + '#_#' + dod_pack_key = '" & row.Item("PLD_ITEM_NO").ToString & "#_#" & _
                '                                                                                            row.Item("pld_pack_key").ToString & "'")


                '            If dtRows.Count > 0 Then
                '                newRow.Item("DO_CODE") = dtRows(0).Item("DO_CODE")
                '                newRow.Item("pad_vnd_code") = dtRows(0).Item("dod_vnd_code")
                '                newRow.Item("pad_carton_no") = dtRows(0).Item("dod_carton_no")

                '                Dim itmString As String = "select d.aitm_pcs_per_pack, " & _
                '                                                "ISNULL(d.aitm_origin, v.aitm_origin) as aitm_origin, " & _
                '                                                "ISNULL(d.aitm_net_weight, v.aitm_net_weight) as aitm_net_weight, " & _
                '                                                "ISNULL(d.aitm_gross_weight, v.aitm_gross_weight) as aitm_gross_weight, " & _
                '                                                "ISNULL(d.aitm_length, v.aitm_length) as aitm_length, " & _
                '                                                "ISNULL(d.aitm_width, v.aitm_width) as aitm_width, " & _
                '                                                "ISNULL(d.aitm_hight, v.aitm_hight) as aitm_hight, " & _
                '                                                "m.itm_pcs_per_uom, " & _
                '                                                "ISNULL(d.aitm_qty_per_ctn, v.aitm_qty_per_ctn) as aitm_qty_per_ctn " & _
                '                                            "from wms_item m " & _
                '                                            "left outer join wms_alt_vend_item d on " & _
                '                                             "m.imp_code = d.imp_code " & _
                '                                                  "and m.storer_code = d.storer_code " & _
                '                                                  "and m.itm_code = d.itm_code " & _
                '                                                  "and m.pack_key = d.pack_key " & _
                '                                                  "and d.vnd_code = '" & gU.dbEncode(dtRows(0).Item("dod_vnd_code").ToString.Trim) & "' " & _
                '                                            "left outer join V_ALT_VEND_ITEM v " & _
                '                                             "on m.IMP_CODE = v.IMP_CODE " & _
                '                                                 "and m.STORER_CODE = v.STORER_CODE " & _
                '                                                 "and m.ITM_CODE = v.ITM_CODE " & _
                '                                                 "and m.PACK_KEY = v.PACK_KEY " & _
                '                                            "where m.itm_code = '" & gU.dbEncode(dtRows(0).Item("dod_itm_code").ToString.Trim) & "' " & _
                '                                            "and m.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                '                                            "and m.imp_code = '" & Session("IMP_CODE") & "' "



                '                Dim itmDt As DataTable = gDB.getDataTable(itmString)


                '                If itmDt.Rows.Count > 0 Then
                '                    newRow.Item("pad_net_weight") = itmDt.Rows(0).Item("aitm_net_weight")
                '                    newRow.Item("pad_gross_weight") = itmDt.Rows(0).Item("aitm_gross_weight")
                '                    newRow.Item("pad_length") = itmDt.Rows(0).Item("aitm_length")
                '                    newRow.Item("pad_width") = itmDt.Rows(0).Item("aitm_width")
                '                    newRow.Item("pad_height") = itmDt.Rows(0).Item("aitm_hight")
                '                    newRow.Item("pad_origin") = itmDt.Rows(0).Item("aitm_origin")
                '                    'newRow.Item("pad_pack_size") = itmDt.Rows(0).Item("aitm_pcs_per_pack")
                '                    newRow.Item("pad_pack_size") = itmDt.Rows(0).Item("itm_pcs_per_uom")
                '                    newRow.Item("PAD_QTY_PER_CTN") = itmDt.Rows(0).Item("aitm_qty_per_ctn")
                '                Else
                '                    newRow.Item("pad_net_weight") = 0
                '                    newRow.Item("pad_gross_weight") = 0
                '                    newRow.Item("pad_length") = 0
                '                    newRow.Item("pad_width") = 0
                '                    newRow.Item("pad_height") = 0
                '                End If
                '            Else
                '                newRow.Item("pad_net_weight") = 0
                '                newRow.Item("pad_gross_weight") = 0
                '                newRow.Item("pad_length") = 0
                '                newRow.Item("pad_width") = 0
                '                newRow.Item("pad_height") = 0
                '            End If

                '            nextSEQ = (CInt(nextSEQ) + 1).ToString

                '            indx += 1

                '            'Session("_M_OB_DO_TMP_pi_seq") = indx

                '            pi_dt.Rows.Add(newRow)
                '        End If
                '    Next

                '    pi_dt.AcceptChanges()
                'End If

                'Session("_M_OB_DO_TMP_pi_dt") = pi_dt


                '==========================================================================================================================================================================
                'No pre-gen of picking and packing ========================================================================================================================================
                '==========================================================================================================================================================================



            End If
        End If
    End Sub

    Protected Sub addItemtoDO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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


            SQLString = "SELECT M.*, D.* " &
            "from WMS_ITEM M, WMS_ALT_VEND_ITEM D " &
            "where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
            "and M.IMP_CODE = '" & Session("IMP_CODE") & "' " &
            "AND M.ITM_CODE + '_000_' + M.PACK_KEY IN ('" & Replace(itemPackList, ", ", "', '") & "') " &
            "AND M.IMP_CODE *= D.IMP_CODE " &
            "AND M.STORER_CODE *= D.STORER_CODE " &
            "AND M.PACK_KEY *= D.PACK_KEY " &
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
                itmKey = dt.Rows(i).Item("pld_item_no").ToString.Trim & "#_#" &
                        dt.Rows(i).Item("pld_pack_key").ToString.Trim & "#_#" &
                        gU.decodeNullOrEmpty(dt.Rows(i).Item("pld_pallet_no").ToString.Trim, "000") & "#_#" &
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

        Dim SQLSTR As String = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " &
                        "cus_area_del, cus_region_del, cus_country_del, cus_cont_per_ord, cus_cont_tel_ord " &
                        "from wms_customer where cus_code = '" & CUS_CODE.SelectedValue & "' " &
                        "and storer_code = '" & STORER_CODE.SelectedValue & "' " &
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

        uiFun.load_dropdown(CUS_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " &
           " storer_code = '" & STORER_CODE.SelectedValue & "' " &
         "and imp_code = '" & Session("IMP_CODE") & "' " &
         "ORDER BY 2", "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

        uiFun.load_dropdown(DO_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT where " &
                                          " storer_code = '" & STORER_CODE.SelectedValue & "' " &
                                        "and imp_code = '" & Session("IMP_CODE") & "' " &
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


        SQLString1 = "SELECT DO_STATUS FROM WMS_DELV_ORDER M " &
                    "WHERE DO_STATUS <> 'POSTED' " &
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString1)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The DO status is New, please post before un-post DO!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The DO status is New, please post before un-post DO!", Session("gLang"))
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
                        itmKey = gU.decodeNull(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" &
                                 gU.decodeNull(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" &
                                 gU.decodeNull(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" &
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
                        SrchStr += "select ILOC_BAL_QTY, Convert(datetime, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(datetime, ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
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


                        itmKey = gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" &
                                 gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" &
                                 gU.decodeNull(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" &
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
                            updateSql = "update WMS_CUST_ORDER_D " &
                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                        "sys_lub = '" & Session("usr_id") & "', " &
                                        "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                        "AND ISNULL(COD_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', '') " &
                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                        Else
                            selectSql = "select COD_SEQ as value " &
                                        "from WMS_CUST_ORDER_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                        "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                        "AND ISNULL(COD_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("PLD_BATCH_NO").ToString.Trim) & "', '') " &
                                        "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                        "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " &
                                        "order by convert(int, COD_SEQ) desc "

                            lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

                            updateSql = "update WMS_CUST_ORDER_D " &
                                        "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                            "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                            "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                        "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                        End If

                        'Update coh_rel_qty for Hold stock logic
                        updateSql = "update wms_cust_order_hold " &
                                    "set COH_REL_QTY = case when ISNULL(COH_REL_QTY,0) - " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " < 0 then 0 else ISNULL(COH_REL_QTY,0) - " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " &
                                        "sys_lub = '" & Session("usr_id") & "', " &
                                        "sys_lud = Getdate() " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                    "and wms_cust_order_hold.coh_status <> 'RELEASE' " &
                                    "and ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) > 0 " &
                                    "and exists (select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                        "where wms_cust_order_hold.imp_code = d.imp_code " &
                                        "and wms_cust_order_hold.storer_code = d.storer_code " &
                                        "and wms_cust_order_hold.co_code = d.co_code " &
                                        "and wms_cust_order_hold.cod_seq = d.cod_seq " &
                                        "and c.imp_code = d.imp_code " &
                                        "and c.storer_code = d.storer_code " &
                                        "and c.co_code = d.co_code " &
                                        "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                                        "and ISNULL(d.COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " &
                                        "and ISNULL(d.COD_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', '') " &
                                        "AND d.COD_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " &
                                        "AND d.COD_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "') "

                        gDB.amendData(updateSql, gConn, transaction)

                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                        REM get Tx Information
                        Dim SrchStr1 As String = ""
                        SrchStr1 = "select IOSX_DRUM_ID,IOSX_DRUM_LEVEL from WMS_OUT_S_TX, WMS_DO_PICKLIST_D WHERE WMS_OUT_S_TX.IO_SYS_SEQ = WMS_DO_PICKLIST_D.PLD_TX_ID " &
                                    "AND WMS_DO_PICKLIST_D.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and WMS_DO_PICKLIST_D.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and WMS_DO_PICKLIST_D.DO_CODE = '" & gU.dbEncode(DO_CODE.Text) & "' " &
                                    "and WMS_DO_PICKLIST_D.PLD_SEQ = '" & gU.dbEncode(rows.Item("PLD_SEQ").ToString.Trim) & "' "
                        Dim txDt As DataTable = gDB.getDataTable(SrchStr1, gConn, transaction)
                        If txDt.Rows.Count > 0 Then
                            st.IOS_DRUM_ID = txDt.Rows(0).Item("IOSX_DRUM_ID").ToString
                            st.IOS_DRUM_LEVEL = txDt.Rows(0).Item("IOSX_DRUM_LEVEL").ToString
                        Else
                            st.IOS_DRUM_ID = ""
                            st.IOS_DRUM_LEVEL = ""
                        End If

                        ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKIN, gConn, transaction, errorMsg)

                        ifPass = st.UnPostSERIAL(StockTrans.IO_TYPE.STOCKIN, gConn, transaction, errorMsg)

                        If Not ifPass Then
                            Exit For
                        End If
                    Next

                    If ifPass Then

                        updateSql = "update WMS_CUST_ORDER " &
                                    "set CO_STATUS = 'CLOSED', " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                    "AND NOT EXISTS (" &
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_CUST_ORDER " &
                                    "set CO_STATUS = 'PARTIAL', " &
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                    "sys_lud = Getdate() " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                    "AND EXISTS (" &
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                        "AND D.COD_POST_QTY > 0) " &
                                    "AND EXISTS (" &
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                        "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                                        "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_CUST_ORDER " &
                                    "set CO_STATUS = 'NEW', " &
                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                    "sys_lud = Getdate() " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and CO_CODE = '" & gU.dbEncode(DO_CO_CODE.Text) & "' " &
                                    "AND NOT EXISTS (" &
                                        "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                        "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                        "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                        "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                        "AND D.COD_POST_QTY > 0) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_DELV_ORDER " &
                                    "set DO_STATUS = 'PICKED', " &
                                        "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                        "sys_lud = Getdate() " &
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
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
                    reloadPage("Record has been Un-Posted successfully!")

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
        Dim add_sql2 As String = ""

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
                add_sql2 = "convert(datetime, '" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & "'," & DDFORMAT & ") "
                add_sql = "Convert(datetime, '" & gU.dbEncode(DO_CONF_DELDATE.Text.Trim) & " " & gU.dbEncode(DO_CONF_DELTIME.Text.Trim) & "', 103) "
            Else
                add_sql2 = "NULL"
                add_sql = "NULL"
            End If


            updateSQL = "Update wms_delv_order set " &
                        "DO_STORER_REM = " & gU.convdbNVCData(gU.dbEncode(DO_STORER_REM.Text.Trim)) & ", " &
                        "DO_CONF_DELDATE=" & add_sql2 & ", " &
                        "DO_CONF_DELTIME=" & add_sql & ", " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
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

                updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='PICKED', sys_LUD=Getdate(), sys_lub='" & gU.dbEncode(Session("usr_id")) & "' " &
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

            updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='NEW', sys_LUD=Getdate(), sys_lub='" & gU.dbEncode(Session("usr_id")) & "' " &
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

    'Protected Sub GridView2_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView2.RowCommand
    '    Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

    '    Select Case e.CommandName
    '        Case "DLVIDEO"
    '            GetDownloadFile(e.CommandArgument, "Video_" & (gvRow.RowIndex + 1))
    '        Case "DELVIDEO"

    '            Deleted_Video_File(CType(gvRow.FindControl("dov_seq"), Label).Text, e.CommandArgument)


    '    End Select
    'End Sub

    'Protected Sub GridView2_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
    '    Select Case e.Row.RowType
    '        Case DataControlRowType.DataRow
    '            CType(e.Row.FindControl("dov_seq"), Label).Text = DataBinder.Eval(e.Row.DataItem, "dov_seq").ToString.Trim
    '            CType(e.Row.FindControl("DOV_WEIGHT"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOV_WEIGHT").ToString.Trim
    '            CType(e.Row.FindControl("DOV_BOX_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOV_BOX_NO").ToString.Trim
    '            'CType(e.Row.FindControl("DOV_VIDEO_NAME"), LinkButton).Text = DataBinder.Eval(e.Row.DataItem, "DOV_VIDEO_NAME").ToString.Trim
    '            CType(e.Row.FindControl("DOV_VIDEO_NAME"), LinkButton).Text = "Video " & (e.Row.RowIndex + 1)
    '            CType(e.Row.FindControl("DOV_VIDEO_NAME"), LinkButton).CommandArgument = DataBinder.Eval(e.Row.DataItem, "DOV_VIDEO_NAME").ToString.Trim

    '            CType(e.Row.FindControl("SYS_CD"), Label).Text = DataBinder.Eval(e.Row.DataItem, "SYS_CD").ToString.Trim
    '            CType(e.Row.FindControl("SYS_CB"), Label).Text = DataBinder.Eval(e.Row.DataItem, "SYS_CB").ToString.Trim

    '            CType(e.Row.FindControl("btnDel"), Button).CommandArgument = DataBinder.Eval(e.Row.DataItem, "DOV_VIDEO_NAME").ToString.Trim
    '            If Session("gLang") = "E" Then
    '                CType(e.Row.FindControl("btnDel"), Button).Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
    '                CType(e.Row.FindControl("btnDel"), Button).Text = "Delete"
    '            ElseIf Session("gLang") = "C" Then
    '                CType(e.Row.FindControl("btnDel"), Button).Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
    '                CType(e.Row.FindControl("btnDel"), Button).Text = "删除"
    '            End If
    '            If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
    '                CType(e.Row.FindControl("lbl_sys_cd"), Label).Visible = False
    '                CType(e.Row.FindControl("lbl_sys_cb"), Label).Visible = False
    '                CType(e.Row.FindControl("SYS_CD"), Label).Visible = False
    '                CType(e.Row.FindControl("SYS_CB"), Label).Visible = False
    '            End If


    '            If DataBinder.Eval(e.Row.DataItem, "DOV_VIDEO_NAME").ToString.Trim <> "" Then
    '                Dim vdPath As String = ResolveUrl("~/" & gU.getConfig("VIDEO_VIRTUAL_FOLDER") & "/" & DataBinder.Eval(e.Row.DataItem, "DOV_VIDEO_NAME").ToString.Trim)
    '                CType(e.Row.FindControl("DOV_VIDEO_NAME_SHOW"), Label).Text = "<video id=""vd_" & CType(e.Row.FindControl("DOV_VIDEO_NAME_SHOW"), Label).ClientID & """ width=""480"" height=""360"" controls > " & _
    '                                        "   <source src=""" & vdPath & """ type=""video/mp4""> " & _
    '                                        "      <object data=""" & vdPath & """ width=""480"" height=""360""  > " & _
    '                                        "         <embed src=""" & vdPath & """ width=""480"" height=""360""  > " & _
    '                                        "      </object> " & _
    '                                        "</video> "
    '            End If

    '            CType(e.Row.FindControl("btnPrint"), Button).OnClientClick = "javascript:openPackLabel('" & Session("DO_CODE") & "','" & DataBinder.Eval(e.Row.DataItem, "dov_seq").ToString.Trim & "');return false;"

    '    End Select
    'End Sub

    Private Sub GetDownloadFile(ByVal SysFilename As String, ByVal UserFileName As String)
        Dim fullFilePath As String = gU.getConfig("FTP_VIDEO_FOLDER") & SysFilename
        Dim uFileName As String = UserFileName & ".mp4"

        If fullFilePath <> "" Then
            Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(fullFilePath)

            If nFile.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" & uFileName)
                Response.AddHeader("Content-Length", nFile.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(nFile.FullName)
                Response.End()
            Else
                ClientScript.RegisterStartupScript(Me.GetType(), "JSFUN", "alert('File Not Exists!');", True)
                'Response.Write("This file does not exist.")
            End If
        End If
    End Sub

    'Public Sub Deleted_Video_File(ByVal dov_seq As String, ByVal FileName As String)
    '    Dim v_sql As String = ""
    '    Dim f_name As String
    '    Dim cnn As SqlConnection
    '    Dim transaction As SqlTransaction

    '    Dim paP As GlobalDBFunc.DBCmdPara
    '    paP = New GlobalDBFunc.DBCmdPara

    '    f_name = FileName

    '    cnn = gDB.getConnection()
    '    transaction = cnn.BeginTransaction()

    '    Try

    '        Dim videoPath As String = gU.getConfig("FTP_VIDEO_FOLDER")
    '        Dim affectedRow As Integer = 0

    '        Dim doCode As String = Session("DO_CODE")

    '        paP = New GlobalDBFunc.DBCmdPara
    '        v_sql = "Delete from WMS_DELV_ORDER_VIDEO " & _
    '                " Where imp_code=" & paP.AP(Session("imp_code")) & " and storer_code=" & paP.AP(Session("STORER_CODE")) & " AND DO_CODE=" & paP.AP(doCode) & _
    '                " AND DOV_SEQ=" & paP.AP(dov_seq)
    '        affectedRow = gDB.amendData(v_sql, cnn, transaction, paP)

    '        If affectedRow > 0 Then
    '            If File.Exists(videoPath & "\" & f_name) Then
    '                File.Delete(videoPath & "\" & f_name)
    '            End If
    '        End If

    '        uiFun.displayMsg(Me, "", "Video File Deleted Successfully!", Session("gLang"))
    '        transaction.Commit()

    '        Call BindGV()

    '    Catch ex As Exception
    '        If Not transaction Is Nothing Then
    '            transaction.Rollback()
    '            transaction = Nothing
    '        End If
    '        'Throw ex
    '        Response.Write(ex.Message)
    '        Response.Write("<BR>" & gDB.getCmdSql(v_sql, paP))
    '    Finally
    '        If cnn IsNot Nothing Then
    '            If cnn.State = ConnectionState.Open Then
    '                cnn.Close()
    '                cnn.Dispose()
    '            End If
    '        End If
    '    End Try

    'End Sub

    'Protected Sub GridView3_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
    '    Select e.Row.RowType
    '        Case DataControlRowType.DataRow
    '            CType(e.Row.FindControl("DOB_FTRACK_NO"), LinkButton).Text = DataBinder.Eval(e.Row.DataItem, "DOB_FTRACK_NO").ToString.Trim
    '            CType(e.Row.FindControl("DOB_FTRACK_NO"), LinkButton).OnClientClick = "javascript:OpenTrack('" & Session("STORER_CODE") & "', '" & DataBinder.Eval(e.Row.DataItem, "DOB_FTRACK_NO").ToString.Trim & "');return false;"
    '    End Select
    'End Sub

    Protected Sub cSBBtn_Click(sender As Object, e As EventArgs) Handles cSBBtn.Click

    End Sub

    Protected Sub DOD_WH_CODE_TextChanged(sender As Object, e As EventArgs)
        Dim txtWHCODE As TextBox = CType(sender, TextBox)
        Dim Row = CType(txtWHCODE.Parent.Parent, GridViewRow)
        Dim nDropDown = CType(Row.FindControl("dod_fl_code"), DropDownList)
        uiFun.load_dropdown(nDropDown, "Select distinct Convert(nvarchar(20),a.FL_NUM) as CODE,a.FL_NAME as NAME from WMS_WH_FL a where a.WH_CODE='" & txtWHCODE.Text & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
    End Sub

    Protected Sub DOD_FL_CODE_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddlFLCODE As DropDownList = CType(sender, DropDownList)
        Dim Row = CType(ddlFLCODE.Parent.Parent, GridViewRow)
        If ddlFLCODE.SelectedValue <> Nothing Then
            Dim txtWHCODE = CType(Row.FindControl("DOD_WH_CODE"), TextBox).Text
            Dim nDropDown = CType(Row.FindControl("DOD_LOC_WH"), DropDownList)
            uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & txtWHCODE & "' and a.FL_NUM='" & ddlFLCODE.SelectedValue.ToString() & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
        End If
    End Sub
    'Generate Picking List'
    Protected Sub btnGenPickingList_Click(sender As Object, e As System.EventArgs) Handles btnGenPickingList.Click
        Dim SQLString As String
        Dim updateSQL As String = ""
        Dim sbCmdText As New StringBuilder
        Dim cmd As New SqlCommand
        Dim codt As DataTable
        Dim coDtl As DataTable
        Dim Qty As Double = 0
        Dim gConn = gDB.getConnection()
        Dim SeqNo As Int16 = 0

        Try
            Dim doCode, storerCode As String
            If Session("DO_CODE") <> "" Then
                doCode = Session("DO_CODE")
                storerCode = Session("STORER_CODE")
            Else
                doCode = Server.UrlDecode(Request("DO_CODE"))
                storerCode = Server.UrlDecode(Request("STORER_CODE"))
                Session("DO_CODE") = doCode
                Session("STORER_CODE") = storerCode
            End If

            SQLString = "select IMP_CODE,STORER_CODE,DOD_WH_CODE,DOD_FL_CODE,DOD_LOC_WH,DO_CODE,DOD_ITM_CODE,DOD_PALLET_NO,DOD_BATCH_NO,DOD_VEND_SEG,DOD_PACK_KEY,DOD_STATUS,DOD_EXPIRY_DATE,DOD_MANU_DATE,DOD_SEQ,DOD_QTY,DOD_SS_QTY,SYS_LUB,SYS_LUD,SYS_CD,SYS_CB from WMS_DELV_ORDER_D WHERE DO_CODE='" + doCode + "'"
            coDtl = gDB.getDataTable(SQLString)
            For Each rowdtl As DataRow In coDtl.Rows
                Dim RdSeq As String
                SQLString = "Select IsNUll(Count(DO_CODE),0)+1 AS PLD_SEQ from WMS_DO_PICKLIST_D where DO_CODE='" + doCode + "' and IMP_CODE='" & rowdtl("IMP_CODE").ToString.Trim() & "' and STORER_CODE='" & (rowdtl("STORER_CODE").ToString.Trim) & "'"
                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString)
                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                    RdSeq = 0
                Else
                    RdSeq = dtRPSD.Rows(0)("PLD_SEQ").ToString
                End If

                Dim PldFL, fl As String
                PldFL = rowdtl("DOD_LOC_WH").ToString.Trim()
                fl = PldFL.Substring(0, 2)

                Dim PldAR, ar As String
                PldAR = rowdtl("DOD_LOC_WH").ToString.Trim()
                ar = PldAR.Substring(2, 2)

                Dim PldRACK, RACK As String
                PldRACK = rowdtl("DOD_LOC_WH").ToString.Trim()
                RACK = PldRACK.Substring(4, 2)

                Dim PldBIN, BIN As String
                PldBIN = rowdtl("DOD_LOC_WH").ToString.Trim()
                BIN = PldBIN.Substring(6, 2)

                SQLString = "SELECT * FROM WMS_DO_PICKLIST_D where DO_CODE='" + doCode + "' and IMP_CODE='" & rowdtl("IMP_CODE").ToString.Trim() & "' and STORER_CODE='" & (rowdtl("STORER_CODE").ToString.Trim) & "' and DOD_SEQ='" & (rowdtl("DOD_SEQ").ToString.Trim) & "'"
                codt = gDB.getDataTable(SQLString)
                If codt Is Nothing Or codt.Rows.Count <= 0 Then
                    SQLString = "Insert into WMS_DO_PICKLIST_D(IMP_CODE,STORER_CODE,DO_CODE,PLD_SEQ,PLD_ITEM_NO,PLD_PACK_KEY,PLD_PALLET_NO,PLD_WH,PLD_LOC,PLD_FLOOR,PLD_BATCH_NO,PLD_DO_QTY,PLD_EXPIRY_DATE,PLD_MANU_DATE,PLD_STATUS,PLD_VEND_SEG,DOD_SEQ,PLD_SS_QTY,PLD_PICKED_BY,PLD_ITEM_QTY,PLD_FOI_QTY,PLD_AREA,PLD_RACK,PLD_BIN,SYS_LUB,SYS_LUD,SYS_CD,SYS_CB) VALUES (@IMP_CODE,@STORER_CODE,@DO_CODE,@PLD_SEQ,@PLD_ITEM_NO,@PLD_PACK_KEY,@PLD_PALLET_NO,@PLD_WH,@PLD_LOC,@PLD_FLOOR,@PLD_BATCH_NO,@PLD_DO_QTY,@PLD_EXPIRY_DATE,@PLD_MANU_DATE,@PLD_STATUS,@PLD_VEND_SEG,@DOD_SEQ,@PLD_SS_QTY,@PLD_PICKED_BY,@PLD_ITEM_QTY,@PLD_FOI_QTY,@PLD_AREA,@PLD_RACK,@PLD_BIN,@SYS_LUB,@SYS_LUD,@SYS_CD,@SYS_CB)"
                    cmd = New SqlCommand(SQLString, gConn)
                    cmd.Parameters.AddWithValue("@IMP_CODE", rowdtl("IMP_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@STORER_CODE", rowdtl("STORER_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@DO_CODE", doCode)
                    cmd.Parameters.AddWithValue("@PLD_SEQ", RdSeq)
                    cmd.Parameters.AddWithValue("@PLD_ITEM_NO", rowdtl("DOD_ITM_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PACK_KEY", rowdtl("DOD_PACK_KEY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PALLET_NO", rowdtl("DOD_PALLET_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_WH", rowdtl("DOD_WH_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_LOC", rowdtl("DOD_LOC_WH").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_FLOOR", fl)
                    cmd.Parameters.AddWithValue("@PLD_BATCH_NO", rowdtl("DOD_BATCH_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_DO_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_EXPIRY_DATE", rowdtl("DOD_EXPIRY_DATE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_MANU_DATE", rowdtl("DOD_MANU_DATE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_STATUS", "PICKED")
                    cmd.Parameters.AddWithValue("@PLD_VEND_SEG", rowdtl("DOD_VEND_SEG").ToString.Trim)
                    cmd.Parameters.AddWithValue("@DOD_SEQ", rowdtl("DOD_SEQ").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_SS_QTY", rowdtl("DOD_SS_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PICKED_BY", Session("usr_id"))
                    cmd.Parameters.AddWithValue("@PLD_ITEM_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_FOI_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_AREA", ar)
                    cmd.Parameters.AddWithValue("@PLD_RACK", RACK)
                    cmd.Parameters.AddWithValue("@PLD_BIN", BIN)
                    cmd.Parameters.AddWithValue("@SYS_LUB", rowdtl("SYS_LUB").ToString.Trim)
                    cmd.Parameters.AddWithValue("@SYS_LUD", rowdtl("SYS_LUD").ToString.Trim)
                    cmd.Parameters.AddWithValue("@SYS_CD", rowdtl("SYS_CD").ToString.Trim)
                    cmd.Parameters.AddWithValue("@SYS_CB", rowdtl("SYS_CB").ToString.Trim)
                    cmd.CommandType = System.Data.CommandType.Text
                    cmd.ExecuteScalar()

                    updateSQL = "Update WMS_DELV_ORDER SET DO_STATUS='PICKED' WHERE DO_CODE='" + doCode + "' and STORER_CODE='" + storerCode + "'"
                    gDB.amendData(updateSQL)

                    uiFun.displayMsg(Me, "", "Generated Picked List successfully!!", Session("gLang"))

                Else
                    'UPDATE
                    sbCmdText = New StringBuilder()
                    sbCmdText.Append("Update WMS_DO_PICKLIST_D Set ")
                    sbCmdText.Append("PLD_ITEM_NO = @PLD_ITEM_NO,")
                    sbCmdText.Append("PLD_PACK_KEY = @PLD_PACK_KEY,")
                    sbCmdText.Append("PLD_PALLET_NO = @PLD_PALLET_NO,")
                    sbCmdText.Append("PLD_WH = @PLD_WH,")
                    sbCmdText.Append("PLD_LOC = @PLD_LOC,")
                    sbCmdText.Append("PLD_FLOOR = @PLD_FLOOR,")
                    sbCmdText.Append("PLD_BATCH_NO = @PLD_BATCH_NO,")
                    sbCmdText.Append("PLD_DO_QTY = @PLD_DO_QTY,")
                    sbCmdText.Append("PLD_EXPIRY_DATE = @PLD_EXPIRY_DATE,")
                    sbCmdText.Append("PLD_MANU_DATE = @PLD_MANU_DATE,")
                    sbCmdText.Append("PLD_STATUS = @PLD_STATUS,")
                    sbCmdText.Append("PLD_VEND_SEG = @PLD_VEND_SEG,")
                    sbCmdText.Append("DOD_SEQ = @DOD_SEQ,")
                    sbCmdText.Append("PLD_SS_QTY = @PLD_SS_QTY,")
                    sbCmdText.Append("PLD_PICKED_BY = @PLD_PICKED_BY,")
                    sbCmdText.Append("PLD_ITEM_QTY = @PLD_ITEM_QTY,")
                    sbCmdText.Append("PLD_FOI_QTY = @PLD_FOI_QTY,")
                    sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                    sbCmdText.Append("SYS_LUD = @SYS_LUD")
                    sbCmdText.Append(" Where IMP_CODE = @IMP_CODE and STORER_CODE=@STORER_CODE and PLD_SEQ=@PLD_SEQ and DO_CODE=@DO_CODE")
                    cmd = New SqlCommand(sbCmdText.ToString(), gConn)
                    cmd.Parameters.AddWithValue("@IMP_CODE", rowdtl("IMP_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@STORER_CODE", rowdtl("STORER_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@DO_CODE", doCode)
                    cmd.Parameters.AddWithValue("@PLD_SEQ", RdSeq)
                    cmd.Parameters.AddWithValue("@PLD_ITEM_NO", rowdtl("DOD_ITM_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PACK_KEY", rowdtl("DOD_PACK_KEY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PALLET_NO", rowdtl("DOD_PALLET_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_WH", rowdtl("DOD_WH_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_LOC", rowdtl("DOD_LOC_WH").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_FLOOR", rowdtl("DOD_FL_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_BATCH_NO", rowdtl("DOD_BATCH_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_DO_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_EXPIRY_DATE", rowdtl("DOD_EXPIRY_DATE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_MANU_DATE", rowdtl("DOD_MANU_DATE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_STATUS", rowdtl("DOD_STATUS").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_VEND_SEG", rowdtl("DOD_VEND_SEG").ToString.Trim)
                    cmd.Parameters.AddWithValue("@DOD_SEQ", rowdtl("DOD_SEQ").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_SS_QTY", rowdtl("DOD_SS_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_PICKED_BY", Session("usr_id"))
                    cmd.Parameters.AddWithValue("@PLD_ITEM_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PLD_FOI_QTY", rowdtl("DOD_QTY").ToString.Trim)
                    cmd.Parameters.AddWithValue("@SYS_LUB", rowdtl("SYS_LUB").ToString.Trim)
                    cmd.Parameters.AddWithValue("@SYS_LUD", rowdtl("SYS_LUD").ToString.Trim)
                    cmd.CommandType = System.Data.CommandType.Text
                    cmd.ExecuteScalar()
                    uiFun.displayMsg(Me, "", "Picked List already Generated!!", Session("gLang"))
                End If

            Next

        Catch ex As Exception
            reloadPage(ex.Message)
        End Try

    End Sub

    Public Sub INSPOSTDATA(STORER_CODE As String, BATCH_NO1 As String, BATCH_NO2 As String, LOCATION1 As String, LOCATION2 As String, WH_CODE As String, ADD_ITM_CODE As String, ADD_REV_QTY1 As Double, ADD_REV_QTY2 As Double, ADD_ORG_QTY1 As Double, ADD_ORG_QTY2 As Double, ADD_VAR_QTY As Double, ADD_EXPIRY_DATE1 As DateTime, ADD_EXPIRY_DATE2 As DateTime, ADD_PALLET_NO1 As String, ADD_PALLET_NO2 As String, transaction As SqlTransaction, gConn As SqlConnection)
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim nextNo As String
        Dim updtSql As String
        'Dim gConn = gDB.getConnection()
        'Dim transaction As SqlTransaction
        'transaction = gConn.BeginTransaction()

        Try
            nextNo = DB.getDocNo("SADJ", gConn, transaction)
            SQLString = "Select * from WMS_STOCK_ADJUST " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and AD_CODE='" & nextNo & "'"
            Dim dtROCode As DataTable
            dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                'INSERT
                SQLString = "Insert into WMS_STOCK_ADJUST(IMP_CODE,STORER_CODE,AD_CODE,AD_DATE,AD_TYPE,AD_STATUS,AD_WH,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB) VALUES (@IMP_CODE, @STORER_CODE,@AD_CODE,@AD_DATE,@AD_TYPE,@AD_STATUS,@AD_WH,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@AD_TYPE", "ADJ")
                cmd.Parameters.AddWithValue("@AD_STATUS", "NEW")
                cmd.Parameters.AddWithValue("@AD_WH", WH_CODE)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                'for details'
                SQLString = "Insert into WMS_STOCK_ADJUST_D(IMP_CODE,STORER_CODE,AD_CODE,AD_SEQ,ADD_PACK_KEY,ADD_ORG_QTY,ADD_REV_QTY,ADD_EXPIRY_DATE,ADD_ITM_CODE,ADD_VAR_QTY,ADD_BATCH_NO,ADD_LOC,ADD_ORG_QTY2,ADD_REV_QTY2,ADD_VAR_QTY2,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB,ADD_PALLET_NO) VALUES (@IMP_CODE,@STORER_CODE,@AD_CODE,@AD_SEQ,@ADD_PACK_KEY,@ADD_ORG_QTY,@ADD_REV_QTY,@ADD_EXPIRY_DATE,@ADD_ITM_CODE,@ADD_VAR_QTY,@ADD_BATCH_NO,@ADD_LOC,@ADD_ORG_QTY2,@ADD_REV_QTY2,@ADD_VAR_QTY2,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB,@ADD_PALLET_NO)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_SEQ", "1")
                cmd.Parameters.AddWithValue("@ADD_PACK_KEY", "1")
                cmd.Parameters.AddWithValue("@ADD_ITM_CODE", ADD_ITM_CODE)
                cmd.Parameters.AddWithValue("@ADD_BATCH_NO", BATCH_NO1)
                cmd.Parameters.AddWithValue("@ADD_LOC", LOCATION1)
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY", ADD_ORG_QTY1)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY", ADD_REV_QTY1)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY", (ADD_REV_QTY1 - ADD_ORG_QTY1))
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_EXPIRY_DATE", ADD_EXPIRY_DATE1)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.Parameters.AddWithValue("@ADD_PALLET_NO", ADD_PALLET_NO1)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                SQLString = "Insert into WMS_STOCK_ADJUST_D(IMP_CODE,STORER_CODE,AD_CODE,AD_SEQ,ADD_PACK_KEY,ADD_ORG_QTY,ADD_REV_QTY,ADD_EXPIRY_DATE,ADD_ITM_CODE,ADD_VAR_QTY,ADD_BATCH_NO,ADD_LOC,ADD_ORG_QTY2,ADD_REV_QTY2,ADD_VAR_QTY2,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB,ADD_PALLET_NO) VALUES (@IMP_CODE,@STORER_CODE,@AD_CODE,@AD_SEQ,@ADD_PACK_KEY,@ADD_ORG_QTY,@ADD_REV_QTY,@ADD_EXPIRY_DATE,@ADD_ITM_CODE,@ADD_VAR_QTY,@ADD_BATCH_NO,@ADD_LOC,@ADD_ORG_QTY2,@ADD_REV_QTY2,@ADD_VAR_QTY2,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB,@ADD_PALLET_NO)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_SEQ", "2")
                cmd.Parameters.AddWithValue("@ADD_PACK_KEY", "1")
                cmd.Parameters.AddWithValue("@ADD_ITM_CODE", ADD_ITM_CODE)
                cmd.Parameters.AddWithValue("@ADD_BATCH_NO", BATCH_NO2)
                cmd.Parameters.AddWithValue("@ADD_LOC", LOCATION2)
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY", ADD_ORG_QTY2)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY", ADD_REV_QTY2)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY", (ADD_REV_QTY2 - ADD_ORG_QTY2))
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_EXPIRY_DATE", ADD_EXPIRY_DATE2)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.Parameters.AddWithValue("@ADD_PALLET_NO", ADD_PALLET_NO2)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()
            End If

            SQLString = "SELECT [IMP_CODE],[STORER_CODE],[AD_CODE],[AD_SEQ],[ADD_ITM_CODE],[ADD_PACK_KEY] " &
",[ADD_LOC],[ADD_ORG_QTY],[ADD_REV_QTY],[ADD_VAR_QTY],[ADD_REM],[ADD_PALLET_NO] " &
",[SYS_LUB],[SYS_LUD],[SYS_CD],[SYS_CB],[ADD_BATCH_NO] " &
",[ADD_VND_CODE],convert(varchar(10),[ADD_EXPIRY_DATE],103) ADD_EXPIRY_DATE,convert(varchar(10),[ADD_MANU_DATE],103)ADD_MANU_DATE,[ADD_ORG_QTY2],[ADD_REV_QTY2] " &
",[ADD_VAR_QTY2],[ADD_SERIAL_NO],[DRUM_ID],[DRUM_LEVEL]FROM[dbo].[WMS_STOCK_ADJUST_D] " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and AD_CODE='" & nextNo & "'"
            Dim dtRODCode As DataTable
            dtRODCode = gDB.getDataTable(SQLString, gConn, transaction)

            If dtRODCode IsNot Nothing AndAlso dtRODCode.Rows.Count > 0 Then
                For Each rows As DataRow In dtRODCode.Rows
                    'FOR POSTING'
                    st.STORER_CODE = STORER_CODE
                    st.ITM_CODE = gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "")
                    st.IO_CUST_CODE = ""
                    st.IO_AREA = ""
                    st.IO_DOC = "SADJ"
                    st.IO_DOC_ID = nextNo
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_WH = WH_CODE
                    st.PALLET_NO = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                    st.IO_LOC = gU.decodeNull(rows.Item("add_loc").ToString.Trim, "")
                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, "")
                    st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, "")
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "")

                    Dim txQty As Double = 0

                    If gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") >
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = 1
                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") <
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)

                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = 1
                            st.UpdateStockSerialTrans("OUT", gConn, transaction)
                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") >
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0")

                        Dim drumDt As DataTable
                        SQLString = "SELECT ILBS_DRUM_ID, ILBS_DRUM_LEVEL FROM WMS_ITEM_LOC_BAL_S WHERE " &
                             "IMP_CODE = 'WMS' " &
                                    "AND STORER_CODE = '" & STORER_CODE & "' " &
                                    "AND ITM_CODE = '" & gU.decodeNull(ADD_ITM_CODE, "") & "' " &
                                    "AND PACK_KEY = '1' " &
                                    "AND ILBS_SERIAL_NO = '" & gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") & "' AND ILBS_QTY2 > 0 "
                        REM **********************

                        drumDt = gDB.getDataTable(SQLString, gConn, transaction)

                        If drumDt.Rows.Count > 0 Then
                            st.IOS_DRUM_ID = drumDt.Rows(0).Item("ILBS_DRUM_ID").ToString
                            st.IOS_DRUM_LEVEL = drumDt.Rows(0).Item("ILBS_DRUM_LEVEL")
                        End If

                        st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                        st.IOS_QTY2 = txQty
                        st.IO_QTY = 1
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)
                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SL = "Y"
                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") <
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0")
                        st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                        st.IOS_QTY2 = txQty
                        st.IO_QTY = 1
                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)
                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SL = "Y"
                            st.UpdateStockSerialTrans("OUT", gConn, transaction)
                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") =
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        'If transaction IsNot Nothing Then
                        '    transaction.Rollback()
                        'End If
                        Exit Sub
                    End If
                Next

            End If

            updtSql = "update wms_stock_adjust " &
                            "set ad_status = 'POSTED' " &
                            "where imp_code = 'WMS' " &
                            "and storer_code = '" & STORER_CODE & "' " &
                            "and ad_code = '" & nextNo & "' "

            gDB.amendData(updtSql, gConn, transaction)

            'For STORER 
            Dim IO_CODE As String
            SQLString = "select Top(1) IO_CODE from EBS_WMS_COMPANY_MASTER where IO_ID = '" & STORER_CODE & "' "
            Dim dtIOCode As DataTable
            dtIOCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtIOCode IsNot Nothing AndAlso dtIOCode.Rows.Count > 0 Then
                IO_CODE = dtIOCode.Rows(0)("IO_CODE").ToString.Trim
            Else
                IO_CODE = ""
            End If

            'For ITEM NUMBER
            Dim ITEM_NUMBER As String
            Dim ITEM_UOM As String
            SQLString = "select Top(1) ITEM_NUMBER,PRIMARY_UOM_CODE from EBS_WMS_ITEM_MASTER where IO_ID = '" & STORER_CODE & "' and ITEM_ID='" & ADD_ITM_CODE & "'  "
            Dim dtITEMNUMBER As DataTable
            dtITEMNUMBER = gDB.getDataTable(SQLString, gConn, transaction)
            If dtITEMNUMBER IsNot Nothing AndAlso dtITEMNUMBER.Rows.Count > 0 Then
                ITEM_NUMBER = dtITEMNUMBER.Rows(0)("ITEM_NUMBER").ToString.Trim
                ITEM_UOM = dtITEMNUMBER.Rows(0)("PRIMARY_UOM_CODE").ToString.Trim
            Else
                ITEM_NUMBER = ""
                ITEM_UOM = ""
            End If

            'FOR SEQ NO.'
            Dim dtADJJSEQ As New DataTable
            SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
            dtADJJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

            'FOR BATCH_NO'
            Dim dtADJBatchNo As New DataTable
            SQLString = "Select REPLACE('WMSLOT'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSLOT1','WMSLOT') as BATCH_NO From WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
            dtADJBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

            'for LOCATION_ID'
            SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + WH_CODE + LOCATION2 + "' and IO_ID='" + STORER_CODE + "'"
            Dim dtLOC As New DataTable
            dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
            Dim LocationId As String = "0"
            If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                LocationId = dtLOC.Rows(0)("LOCID").ToString()
            End If

            'for LOCATOR'
            Dim LOC As String
            LOC = LOCATION2
            Dim location As String = LOC.Substring(0, 2) + "." + LOC.Substring(2, 2) + "." + LOC.Substring(4, 2) + "." + LOC.Substring(6, 2)

            'for TRANSACTION_ID as use seq_no'
            Dim dtSeq As New DataTable
            SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
            dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

            'CompanyId
            Dim CompanyId As Integer
            Dim dtCompanyData As New DataTable
            SQLString = "select Top(1) Isnull(COMPANY_ID,0) COMPANY_ID from EBS_WMS_COMPANY_MASTER where IO_ID=" & STORER_CODE & ""
            dtCompanyData = gDB.getDataTable(SQLString, gConn, transaction)
            If (dtCompanyData IsNot Nothing AndAlso dtCompanyData.Rows.Count > 0) Then
                CompanyId = dtCompanyData.Rows(0)("COMPANY_ID")
            Else
                CompanyId = 0
            End If
            'for WMS_EBS_TRANS_ITX_ACTION'
            'INSERT
            SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,IO_ID,DOC_TYPE,STATUS,BATCH_NO,COMPANY_ID,CREATION_DATE,LAST_UPDATE_DATE) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@IO_ID,@DOC_TYPE,@STATUS,@BATCH_NO,@COMPANY_ID,@CREATION_DATE,@LAST_UPDATE_DATE)"
            cmd = New SqlCommand(SQLString, gConn, transaction)
            cmd.Parameters.AddWithValue("@SOURCE", "WMS")
            cmd.Parameters.AddWithValue("@ACTION", "NEW")
            cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
            cmd.Parameters.AddWithValue("@IO_ID", STORER_CODE)
            cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
            cmd.Parameters.AddWithValue("@COMPANY_ID", CompanyId)
            cmd.Parameters.AddWithValue("@STATUS", "NEW")
            cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
            cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
            cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
            cmd.CommandType = System.Data.CommandType.Text
            cmd.ExecuteScalar()

            SQLString = "Select * from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT " &
                      "WHERE TRANSACTION_ID ='" + dtSeq.Rows(0)("TRANSACTION_ID").ToString + "'"
            Dim dtSTKADJCode As DataTable
            dtSTKADJCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtSTKADJCode Is Nothing Or dtSTKADJCode.Rows.Count <= 0 Then
                'for Issue'
                'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,SUBINVENTORY_CODE,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@SUBINVENTORY_CODE,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY)"
                'UOM_CODE,@UOM_CODE
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@SEQ_NO", dtADJJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "I")
                cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
                cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", WH_CODE)
                cmd.Parameters.AddWithValue("@LOCATOR", location)
                cmd.Parameters.AddWithValue("@LOCATION_ID", LocationId)
                cmd.Parameters.AddWithValue("@LOT_NUMBER", BATCH_NO2)
                cmd.Parameters.AddWithValue("@ITEM_NUMBER", ITEM_NUMBER)
                cmd.Parameters.AddWithValue("@QUANTITY", ADD_VAR_QTY)
                cmd.Parameters.AddWithValue("@UOM_CODE", ITEM_UOM)
                cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_id"))
                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", DBNull.Value)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                'for Receive'
                'FOR SEQ NO.'
                Dim dtADJSEQ As New DataTable
                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
                dtADJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

                'for LOCATOR'
                Dim RCVLOC As String
                RCVLOC = LOCATION1
                Dim RCVlocation As String = RCVLOC.Substring(0, 2) + "." + RCVLOC.Substring(2, 2) + "." + RCVLOC.Substring(4, 2) + "." + RCVLOC.Substring(6, 2)

                'for LOCATION_ID'
                SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + WH_CODE + LOCATION1 + "' and IO_ID='" + STORER_CODE + "'"
                Dim dtRCVLOC As New DataTable
                dtRCVLOC = gDB.getDataTable(SQLString, gConn, transaction)
                Dim RCVLocationId As String = "0"
                If dtRCVLOC IsNot Nothing AndAlso dtRCVLOC.Rows.Count > 0 Then
                    RCVLocationId = dtRCVLOC.Rows(0)("LOCID").ToString()
                End If

                'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,SUBINVENTORY_CODE,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY,EXPIRY_DATE) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@SUBINVENTORY_CODE,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY,@EXPIRY_DATE)"
                'UOM_CODE,@UOM_CODE
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@SEQ_NO", dtADJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "R")
                cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
                cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", WH_CODE)
                cmd.Parameters.AddWithValue("@LOCATOR", RCVlocation)
                cmd.Parameters.AddWithValue("@LOCATION_ID", RCVLocationId)
                cmd.Parameters.AddWithValue("@LOT_NUMBER", BATCH_NO1)
                cmd.Parameters.AddWithValue("@ITEM_NUMBER", ITEM_NUMBER)
                cmd.Parameters.AddWithValue("@QUANTITY", ADD_VAR_QTY)
                cmd.Parameters.AddWithValue("@UOM_CODE", ITEM_UOM)
                cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_id"))
                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", DBNull.Value)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                cmd.Parameters.AddWithValue("@EXPIRY_DATE", ADD_EXPIRY_DATE1)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()
            End If


        Catch ex As Exception
            Throw ex

        End Try

    End Sub
End Class
