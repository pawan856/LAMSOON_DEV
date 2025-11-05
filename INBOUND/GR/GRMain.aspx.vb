Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Globalization

Partial Class GRMain
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
    Private sn_dt As New DataTable
    Private dtWh, dtUOM, dtUOM2, dtRejReason, dtDtlStatus As DataTable
    'Protected isClickPutAway As Boolean = False

    Private gvCol() As String = {"CB_SELECT", "BTNSPLIT", "BTNHIGHLIGHT", "GRD_DISP_SEQ", "GRD_VND_CODE", "GRD_ITM_CODE",
                                 "ITM_SKU_NO", "GRD_PACK_KEY", "GRD_ITM_NAME", "GRD_PACK", "GRD_PACK_NO", "GRD_PO_QTY", "GRD_RCV_QTY",
                                 "GRD_OS_QTY", "GRD_PALLET_NO", "GRD_NO_OF_CARTON", "GRD_PCS_PER_CARTON", "GRD_CARTON_NO", "GRD_BATCH_NO",
                                 "GRD_REF_NO", "GRD_BRAND", "GRD_SERIES", "GRD_MODEL", "GRD_UOM", "GRD_LENGTH", "GRD_KG",
                                 "SUB_GRD_KG", "GRD_CBM", "GRD_DESTINATION", "GRD_STOCKTYPE", "GRD_PCS_PER_UOM", "GRD_TOT_PCS", "GRD_EXPIRY_DATE", "GRD_MANU_DATE",
                                 "GRD_REJ_QTY", "GRD_REJ_REASON", "GRD_INSP_REQ", "GRD_STATUS", "GRD_UOM2", "GRD_QTY2", "GRD_ON_BEHALF", "GRD_WH"
                                 }

    Private wmsFun As New WMSFunc

    Const dmgLocValue As String = "Damage Zone"

    Dim SubTotalPCS As Double = 0.0
    Dim TotalCar As Double = 0.0
    Dim TotalKG As Double = 0.0
    Dim TotalCBM As Double = 0.0

    Dim firstROClientId As String = ""
    Dim firstOSClientId As String = ""
    Dim DDFORMAT As String = gU.getConfig("DDFORMATNO")

    Structure itmStruct
        Dim itmName As String
        Dim qty As Double
        Dim itmSeq As String
        Dim batchNo As String
        Dim refNo As String
        Dim skuNo As String
        Dim expDate As String
        Dim manuDate As String
        Dim whCode As String
    End Structure

    Structure itmLocQty
        Dim loc As String
        Dim qty As Double
    End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM ****************************
        REM Modify Access Right Here

        'If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
        '    Me.ClientScript.RegisterStartupScript(Me.GetType, "loadcss", "toLoadCSS();", True)
        'End If

        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)

        'moduleAction = Request("moduleAction")

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        'DDFORMAT = gU.getConfig("DDFORMAT")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("ISCOPYING") = "N"

            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(PRJ_CODE, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(GR_TRANS_TYPE, "Select COLC_CODE, COLC_ENG_VALUE from wms_col_code where COLC_TABCOL='WMS_GOODSRCV.GR_TRANS_TYPE' order by COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))
            'uiFun.load_dropdown(VND_CODE, "select VND_CODE, VND_NAME from WMS_VENDOR ORDER BY 2", "VND_CODE", "VND_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(GR_CAT, "Select PO_CATEGORY_CODE,PO_CATEGORY_DESC from PO_CATEGORY_MASTER ORDER BY 2", "PO_CATEGORY_CODE", "PO_CATEGORY_DESC", , Session("gSelectLabel"))

        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            btnRefresh.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            If GR_DATE.Text = "" Then
                GR_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If

            GR_RCV_BY.Text = Session("usr_id")
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Goods Receiving Maintenance"
            lbl_ImageHd.Text = "Goods Receiving Items"
            lbl_GR_CODE.Text = "Receipt No:"
            lbl_GR_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_GR_DATE.Text = "Receive Date:"
            lbl_GR_RCV_BY.Text = "Received By:"
            lbl_GR_WH_CODE.Text = "Subinventory:"
            lbl_PRJ_CODE.Text = "Project:"
            lbl_GR_DOC_TYPE.Text = "Doc. Type:"
            lbl_GR_DOC_NO.Text = "Doc. Code:"
            lbl_PO_DATE.Text = "Doc. Date:"
            lbl_VND_CODE.Text = "NPO Ref Number:"
            lbl_VND_NAME.Text = "Vendor Name:"
            lbl_GR_VND_DNREF.Text = "Deli. Note No.:"
            lbl_GR_DESTINATION.Text = "Destination:"
            lbl_GR_TOT_PALLET.Text = "Total Pallet:"
            lbl_GR_REM.Text = "Remarks:"
            'lbl_GR_TRACK_NO.Text = "Tracking No.:"
            'lbl_GR_REF_NO.Text = "Ref. No.:"
            lbl_GR_EDI_PO_NO.Text = "PO No."
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            lbl_GR_TRANS_TYPE.Text = "Nature of Transaction:"
            lbl_gr_cat.Text = "Category:"
            'NEW FIELD
            lbl_GR_HAWB.Text = "HAWB"
            lbl_GR_INV_NO.Text = "Invoice No."

            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            btnPutAway.Text = "Put Away"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "if(!confirm(""Are you sure to cancel this record?"")) return false;"
            saveBtn1.OnClientClick = "if (!confirm(""Are you sure to save this record?"")) return false;"
            saveBtn2.OnClientClick = "if (!confirm(""Are you sure to save this record?"")) return false;"

            'btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting"")){getLoad();}else{return false;}"
            btnPost.OnClientClick = "if(this.disabled) return false; if(confirm('Are you sure to post this record?\\r\\n(Please save your work before Posting)')) {setTimeout(function() {document.getElementById('btnPost').disabled=true;}, 50); return true;} else {return false;}"
            btnUnPost.OnClientClick = "if (!confirm(""Are you sure to un-post this record?\r\n(Please save your work before Un-Posting)"")) return false;"

            btnNOTE.Text = "Notes"
            'btnItemDelete.OnClientClick = "return confirm(""Are you sure you want to delete selected items?\r\n(Please reset Put Away after deletion)"");"
            If Session("pagemode") = "N" Then
                GR_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lbl_GR_CODE.Text = "收貨編號:"
            lbl_ImageHd.Text = "收貨物件:"
            lbl_GR_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_GR_DATE.Text = "收貨日期:"
            lbl_GR_RCV_BY.Text = "收貨者:"
            lbl_GR_WH_CODE.Text = "子庫存:"
            lbl_PRJ_CODE.Text = "項目:"
            lbl_GR_DOC_TYPE.Text = "文件類型:"
            lbl_GR_DOC_NO.Text = "文件編號:"
            lbl_PO_DATE.Text = "文件日期:"
            lbl_VND_CODE.Text = "NPO 参考编号:"
            lbl_VND_NAME.Text = "供應商名稱:"
            lbl_GR_VND_DNREF.Text = "送貨單號:"
            lbl_GR_DESTINATION.Text = "目的地:"
            lbl_GR_TOT_PALLET.Text = "貨板總量:"
            lbl_GR_REM.Text = "備註:"
            'lbl_GR_TRACK_NO.Text = "追查編號:"
            'lbl_GR_REF_NO.Text = "文件編號:"
            lbl_GR_EDI_PO_NO.Text = "單號:"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            CancelBtn.Text = "取消"
            btnPutAway.Text = "上架"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            selectItemBtn.Text = "選擇物料"
            lbl_GR_TRANS_TYPE.Text = "交易性質:"
            lbl_gr_cat.Text = "类别:"
            btnRefresh.Value = "重新整理"
            btnInsp.Text = "驗貨"
            cSBBtn.Text = "檢查貨品存庫"
            btnPost.Text = "入庫"
            btnUnPost.Text = "反入庫"
            btnNOTE.Text = "工單"
            'New Field
            lbl_GR_HAWB.Text = "航空分運單:"
            lbl_GR_INV_NO.Text = "單據編號:"
            'CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            'saveBtn1.OnClientClick = "return confirm(""確定保存資料?"");"
            'saveBtn2.OnClientClick = "return confirm(""確定保存資料?"");"
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            'btnUnPost.OnClientClick = "return confirm(""確定取消發布資料?"");"

            'btnItemDelete.OnClientClick = "return confirm(""你是否確定要刪除選取了的資料?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[号码会自动产生]"
            'End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        GR_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            STORER_CODE.CssClass = "REQUIRED"
        Else
            STORER_CODE.Enabled = False
        End If

        If Session("pagemode") = "N" Then
            If STORER_CODE.SelectedValue <> "" Then
                ViewState("STORER_CODE") = STORER_CODE.SelectedValue
            End If
        Else
            If ViewState("STORER_CODE") = "" Then
                ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))
            End If
        End If

        If ViewState("STORER_CODE") <> "" Then
            Dim selectSql As String

            selectSql = "select sto_batch_field_ref as VALUE " & _
                        "from WMS_STORER " & _
                        "where STORER_CODE = '" & gU.dbEncode(ViewState("STORER_CODE")) & "' " & _
                        "and IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "'"


            ViewState("STO_BATCH_FIELD_REF") = DB.getValueFromSQL(selectSql)
        End If

        If Not IsPostBack Then
            Session("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("GR_CODE") = ""

            ViewState("GR_CODE") = Server.UrlDecode(Request("GR_CODE"))

            Call BindGV()
        Else
            dt = Session("dt")

            If editMode.Value = "V" Then
                GridView1.DataSource = dt
                GridView1.DataBind()
            End If

            'If Not Session("sn_dt") Is Nothing Then
            '    sn_dt = Session("sn_dt")
            'Else
            '    Dim sStr As String = "select wms_goodsrcv_d_s.*, 'U' as mFlag " & _
            '      "from wms_goodsrcv_d_s, wms_goodsrcv_d d " & _
            '      "where " & _
            '      "d.imp_code = wms_goodsrcv_d_s.imp_code " & _
            '      "and d.storer_code = wms_goodsrcv_d_s.storer_code " & _
            '      "and d.gr_code = wms_goodsrcv_d_s.gr_code " & _
            '      "and d.grd_itm_code = wms_goodsrcv_d_s.grs_itm_code " & _
            '      "and d.gr_code = '" & gU.dbEncode(ViewState("GR_CODE")) & "' " & _
            '      "and d.storer_code = '" & gU.dbEncode(ViewState("STORER_CODE")) & "' " & _
            '      "and d.imp_code = '" & Session("imp_code") & "'"

            '    sStr = sStr & " order by wms_goodsrcv_d_s.grs_disp_seq"

            '    sn_dt = gDB.getDataTable(sStr)

            '    Session("sn_dt") = sn_dt
            'End If
        End If

        'Set Access Right
        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            If Session("usr_pref_storer") <> STORER_CODE.Text Then
                Response.End()
            End If
        End If

        REM ****************************************************************

        If moduleAction.Value = "SELECTRO" Then
            addROtoGR()
        End If

        If moduleAction.Value = "SELECTIM" Then
            addItemtoGR()
        End If

        If moduleAction.Value = "CHECKBARCODE" Then
            checkBarCode()
        End If

        If moduleAction.Value = "SHOW_SERIAL" Then
            showSerial()
        End If

        If GR_DOC_NO.Value <> "" Then
            newrow.Visible = False
        End If

        'ar.addColDef("GRD_DISP_SEQ", "grd_disp_seq", 0)
        'ar.addColDef("GRD_ITM_CODE", "grd_itm_code", 1)
        'ar.addColDef("GRD_PACK_KEY", "grd_pack_key", 2)
        'ar.addColDef("GRD_ITM_NAME", "grd_itm_name", 3)
        'ar.addColDef("GRD_PACK", "grd_pack", 4)
        'ar.addColDef("GRD_PACK_NO", "grd_pack_no", 5)
        'ar.addColDef("GRD_PO_QTY", "grd_po_qty", 6)
        'ar.addColDef("GRD_RCV_QTY", "grd_rcv_qty", 7)
        'ar.addColDef("GRD_OS_QTY", "grd_os_qty", 8)
        ''ar.addColDef("GRD_WH", "grd_wh", 9)
        ''ar.addColDef("GRD_LOC", "grd_loc", 10)
        'ar.addColDef("GRD_PALLET_NO", "grd_pallet_no", 9)
        'ar.addColDef("GRD_CARTON_NO", "grd_carton_no", 10)
        'ar.addColDef("GRD_BATCH_NO", "grd_batch_no", 11)
        'ar.addColDef("GRD_REF_NO", "grd_ref_no", 12)
        'ar.addColDef("GRD_BRAND", "grd_brand", 13)
        'ar.addColDef("GRD_SERIES", "grd_series", 14)
        'ar.addColDef("GRD_MODEL", "grd_model", 15)
        'ar.addColDef("GRD_UOM", "grd_uom", 16)
        'ar.addColDef("GRD_LENGTH", "grd_length", 17)
        'ar.addColDef("GRD_WIDTH", "grd_width", 18)
        'ar.addColDef("GRD_HEIGHT", "grd_height", 19)
        'ar.addColDef("GRD_KG", "grd_kg", 18)
        'ar.addColDef("GRD_CBM", "grd_cbm", 19)
        'ar.addColDef("GRD_DESTINATION", "grd_destination", 20)
        'ar.addColDef("GRD_STOCKTYPE", "grd_stocktype", 21)
        'ar.addColDef("GRD_PCS_PER_UOM", "grd_pcs_per_uom", 22)
        'ar.addColDef("GRD_TOT_PCS", "grd_tot_pcs", 23)

        changeLabel()

        Dim colIdx_StartWith As Integer = 3

        ar.addColDef("GRD_DISP_SEQ", "grd_disp_seq", colIdx_StartWith)
        ar.addColDef("GRD_VND_CODE", "grd_vnd_code", colIdx_StartWith)
        ar.addColDef("GRD_ITM_CODE", "grd_itm_code", colIdx_StartWith)
        ar.addColDef("GRD_PACK_KEY", "grd_pack_key", colIdx_StartWith)
        ar.addColDef("GRD_ITM_NAME", "grd_itm_name", colIdx_StartWith)
        ar.addColDef("GRD_PACK", "grd_pack", colIdx_StartWith)
        ar.addColDef("GRD_PACK_NO", "grd_pack_no", colIdx_StartWith)
        ar.addColDef("GRD_PO_QTY", "grd_po_qty", colIdx_StartWith)
        ar.addColDef("GRD_RCV_QTY", "grd_rcv_qty", colIdx_StartWith)
        ar.addColDef("GRD_OS_QTY", "grd_os_qty", colIdx_StartWith)
        ar.addColDef("GRD_PALLET_NO", "grd_pallet_no", colIdx_StartWith)
        ar.addColDef("GRD_NO_OF_CARTON", "grd_no_of_carton", colIdx_StartWith)
        ar.addColDef("GRD_PCS_PER_CARTON", "grd_pcs_per_carton", colIdx_StartWith)
        ar.addColDef("GRD_CARTON_NO", "grd_carton_no", colIdx_StartWith)
        'ar.addColDef("GRD_LOT_NO", "grd_lot_no", colIdx_StartWith)
        ar.addColDef("GRD_BATCH_NO", "grd_batch_no", colIdx_StartWith)

        ar.addColDef("GRD_REF_NO", "grd_ref_no", colIdx_StartWith)
        'ar.addColDef("GRD_BRAND", "grd_brand", colIdx_StartWith)
        'ar.addColDef("GRD_SERIES", "grd_series", colIdx_StartWith)
        'ar.addColDef("GRD_MODEL", "grd_model", colIdx_StartWith)
        ar.addColDef("GRD_UOM", "grd_uom", colIdx_StartWith)
        ar.addColDef("GRD_LENGTH", "grd_length", colIdx_StartWith)
        ar.addColDef("GRD_WIDTH", "grd_width", colIdx_StartWith, , False)
        ar.addColDef("GRD_HEIGHT", "grd_height", colIdx_StartWith, , False)
        ar.addColDef("GRD_KG", "grd_kg", colIdx_StartWith)
        ar.addColDef("", "", colIdx_StartWith)
        ar.addColDef("GRD_CBM", "grd_cbm", colIdx_StartWith)
        ar.addColDef("GRD_DESTINATION", "grd_destination", colIdx_StartWith)
        ar.addColDef("GRD_STOCKTYPE", "grd_stocktype", colIdx_StartWith)
        ar.addColDef("GRD_PCS_PER_UOM", "grd_pcs_per_uom", colIdx_StartWith)
        ar.addColDef("GRD_TOT_PCS", "grd_tot_pcs", colIdx_StartWith)

        ar.addColDef("GRD_EXPIRY_DATE", "grd_expiry_date", colIdx_StartWith)
        ar.addColDef("GRD_MANU_DATE", "grd_manu_date", colIdx_StartWith)

        ar.addColDef("GRD_REJ_QTY", "grd_rej_qty", colIdx_StartWith)

        ar.addColDef("GRD_REJ_REASON", "grd_rej_reason", colIdx_StartWith)
        ar.addColDef("GRD_INSP_REQ", "grd_insp_req", colIdx_StartWith)

        ar.addColDef("GRD_STATUS", "grd_status", colIdx_StartWith)

        ar.addColDef("GRD_UOM2", "grd_uom2", colIdx_StartWith)
        ar.addColDef("GRD_QTY2", "grd_qty2", colIdx_StartWith)

        'ar.addColDef("GRD_ON_BEHALF", "grd_on_behalf", colIdx_StartWith)
        ar.addColDef("GRD_WH", "grd_wh", colIdx_StartWith)

        REM Generate Common Menu
        cm = New CommonMenu("GR", lheader.Value, GR_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "N"
        cm.haveAttachments = "N"
        cm.haveNotes = "N"
        cm.haveTasks = "N"
        cm.haveEmail = "N"
        cm.haveHistory = "N"
        cm.genCM(cmBar)

        If dt.Rows.Count > 0 Then
            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
        End If

        REM Select RO button
        If GR_DOC_NO.Value <> "" Then
            selectROBtn.Visible = False
            'selectItemBtn.Visible = False
        Else
            'selectROBtn.Attributes.Add("onclick", "window.open('../../cms_search.aspx?menu_code=LOOKUP_RO&pForm=forms[0]&pFunc=selectedRO()&pItemList=" & txt_GR_DOC_NO.ClientID & "|1|L, " & GR_DOC_NO.ClientID & "|1&sc='+document.myform." & STORER_CODE.ClientID & ".value);")
            selectROBtn.Attributes.Add("onclick", "ROLookUp('" & txt_GR_DOC_NO.ClientID & "', '" & GR_DOC_NO.ClientID & "', document.myform." & STORER_CODE.ClientID & ".value);")
        End If

        REM Select RO button
        If GR_STATUS.Text = "POSTED" Then
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Enabled = True
        Else
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")
            cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);")
            cSBBtn.Enabled = True
        End If

        REM ****************************************************************


        If GR_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            'ar.sec_read = "Y"
            CancelBtn.Visible = False
            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue
        ElseIf GR_STATUS.Text = "POSTED" Then
            'ar.sec_viewMode = "Y"
            'ar.sec_write = "N"
            ar.sec_viewMode = "Y"
            btnPost.Visible = False
            btnUnPost.Visible = True
            ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue
        Else
            btnUnPost.Visible = False
        End If


        Dim exceptionEditList As List(Of String)

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("btnCloseDetailSerial")

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
        'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "GR", "WMS_GOODSRCV_D")

        REM ************************
        REM check right ADMIN_GP
        'ADMIN_GP
        If GR_STATUS.Text = "POSTED" Then
            'Dim checkGP As String = "select usr_id from wms_user_group_alloc where grp_code = 'ADMIN_GP' and usr_id = '" & Session("usr_id") & "'"
            'Dim checkGPDt As DataTable
            'checkGPDt = gDB.getDataTable(checkGP)
            'If checkGPDt.Rows.Count > 0 Then
            '    'btnUnPost.Enabled = True
            '    btnUnPost.Enabled = ar.hasBtnRight("BT_GR_UNPOST")
            'Else
            '    btnUnPost.Enabled = False
            'End If
            btnUnPost.Enabled = ar.hasBtnRight("BT_GR_UNPOST")
        End If



        If moduleAction.Value = "SAVEOK" Then
            save()
        End If

        REM generate Link Bar
        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('IB_GR','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("GR_CODE") & "','N');")

        'ar.setFieldCustomize(Me, "IB_GR", STORER_CODE.SelectedValue, "WMS_GOODSRCV")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "IB_GR", STORER_CODE.SelectedValue, "WMS_GOODSRCV_D", gvCol)
        'End If

        Dim size_index As Integer = -1
        Dim kg_index As Integer = -1

        ViewState("size_index") = size_index
        ViewState("kg_index") = kg_index

        If GridView1.Rows.Count > 0 Then
            For v As Integer = 0 To GridView1.Rows(0).Cells.Count - 1
                If GridView1.Rows(0).Cells(v).Visible Then

                    size_index += 1
                    kg_index += 1

                    If GridView1.Columns(v).AccessibleHeaderText = "grd_size" Then
                        ViewState("size_index") = size_index
                    ElseIf GridView1.Columns(v).AccessibleHeaderText = "grd_kg_c" Then
                        ViewState("kg_index") = kg_index
                    End If

                End If
            Next
        End If


        Dim cIndex As Integer = 0
        Dim carIndex As Integer = 0
        Dim qtyIndex As Integer = 0
        Dim kgIndex As Integer = 0
        Dim cbmIndex As Integer = 0

        Dim tmpIdx As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For v As Integer = 0 To GridView1.Rows(0).Cells.Count - 1
                If GridView1.Rows(0).Cells(v).Visible Then
                    Select Case GridView1.Columns(v).AccessibleHeaderText
                        Case "grd_no_of_carton"
                            carIndex = tmpIdx
                        Case "grd_rcv_qty"
                            qtyIndex = tmpIdx
                        Case "grd_kg"
                            kgIndex = tmpIdx
                        Case "grd_cbm"
                            cbmIndex = tmpIdx
                        Case "grd_tot_pcs"
                            cIndex = tmpIdx

                    End Select

                    tmpIdx += 1
                End If
            Next
        End If

        ViewState("cIndex") = cIndex
        ViewState("carIndex") = carIndex
        ViewState("qtyIndex") = qtyIndex
        ViewState("kgIndex") = kgIndex
        ViewState("cbmIndex") = cbmIndex

    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(saveBtn1)
        sm.RegisterAsyncPostBackControl(saveBtn2)
        sm.RegisterAsyncPostBackControl(selectROBtn)
        sm.RegisterAsyncPostBackControl(CancelBtn)
        sm.RegisterAsyncPostBackControl(btnPutAway)
        sm.RegisterAsyncPostBackControl(cSBBtn)
        sm.RegisterAsyncPostBackControl(btnPost)
        sm.RegisterAsyncPostBackControl(btnUnPost)
        sm.RegisterAsyncPostBackControl(btnAddDetailSerial)
        sm.RegisterAsyncPostBackControl(btnDetailSerialOK)
        sm.RegisterAsyncPostBackControl(btnInTransit)
        sm.RegisterAsyncPostBackControl(btnRelease)
        sm.RegisterAsyncPostBackControl(btnInsp)

        For i = 0 To GridView1.Rows.Count - 1
            sm.RegisterAsyncPostBackControl(CType(GridView1.Rows(i).FindControl("btnItemSerial"), ImageButton))
        Next

        'For i = 0 To gvDetailSerialList.Rows.Count - 1
        '    sm.RegisterAsyncPostBackControl(CType(gvDetailSerialList.Rows(i).FindControl("Image_Loc_LookUp"), ImageButton))
        'Next

        If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
            'Set labels, attributes, etc
            'setGeneralControls()

            'Set field access (hide, readonly, view mode, etc)
            'setPageCtrlAccess()
        End If

        moduleAction.Value = ""
    End Sub

    'Private Sub setGeneralControls()
    '    Dim hasOnBehalf As Boolean

    '    hasOnBehalf = False

    '    For i = 0 To GridView1.Rows.Count - 1
    '        If CType(GridView1.Rows(i).FindControl("GRD_ON_BEHALF"), HiddenField).Value = "Y" AndAlso CType(GridView1.Rows(i).FindControl("GRD_STATUS"), DropDownList).SelectedValue <> "INTRANS" Then
    '            hasOnBehalf = True
    '            Exit For
    '        End If
    '    Next

    '    If hasOnBehalf Then
    '        btnInTransit.Visible = True
    '        btnInTransit.OnClientClick = "return confirm('Confirm goods are in Transit to the second warehouse (will generate a new WPO)?');"
    '    End If

    '    If GR_CODE_HF.Value <> "" AndAlso GR_STATUS.Text = "NEW" Then
    '        btnRelease.Visible = True
    '        btnRelease.OnClientClick = "return confirm('Confirm to release Put Away to PDA?');"
    '    End If
    'End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False

        If ctl.ID = "ser_table" Then
            customizectrl = True
        End If
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Or ctl.ID = "btnPutAway" Or ctl.ID = "btnUnPost" Or ctl.ID = "btnInsp" Then
            If ctl.ID = "btnUnpost" Then
                CType(ctl, Button).Enabled = ar.hasBtnRight("BT_GR_UNPOST")
            Else
                CType(ctl, Button).Enabled = True
            End If

            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        Dim selectSql As String

        selectSql = "select wh_code as code, wh_code as name from wms_warehouse " & _
                    "order by wh_code "

        dtWh = gDB.getDataTable(selectSql)

        selectSql = "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1"

        dtUOM = gDB.getDataTable(selectSql)

        selectSql = "select colc_code as code, colc_eng_value as name from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_UOM2' order by COLC_DISPLAY_SEQ"

        dtUOM2 = gDB.getDataTable(selectSql)

        selectSql = "select colc_code as code, colc_eng_value as name " & _
                    "from wms_col_code " & _
                    "where colc_tabcol = 'WMS_GOODSRCV_D.GRD_REJ_REASON' " & _
                    "order by colc_eng_value "

        dtRejReason = gDB.getDataTable(selectSql)

        selectSql = "select colc_code as code, colc_eng_value as name from wms_col_code where colc_tabcol = 'WMS_GOODSRCV_D.GRD_STATUS' order by COLC_DISPLAY_SEQ"

        dtDtlStatus = gDB.getDataTable(selectSql)

    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        'Dim scriptStr As String = ""
        'Dim scriptStr2 As String = ""
        'Dim scriptStr3 As String = ""
        'Dim scriptStr4 As String = ""

        'For Each row In GridView1.Rows
        '    scriptStr += "grd_batch.push('" & CType(row.FindControl("grd_batch_no"), Label).ClientID & "');"
        '    scriptStr2 += "grd_qty.push('" & CType(row.FindControl("grd_rcv_qty"), TextBox).ClientID & "');"
        '    'scriptStr3 += "grd_dc.push('" & CType(row.FindControl("GRD_BATCH_CD"), TextBox).ClientID & "');"
        '    'scriptStr4 += "grd_lot.push('" & CType(row.FindControl("GRD_LOT_NO"), TextBox).ClientID & "');"
        'Next

        'Me.ClientScript.RegisterStartupScript(Me.GetType, "batch", scriptStr, True)
        'Me.ClientScript.RegisterStartupScript(Me.GetType, "qty", scriptStr2, True)
        ''Me.ClientScript.RegisterStartupScript(Me.GetType, "batchCD", scriptStr3, True)
        ''Me.ClientScript.RegisterStartupScript(Me.GetType, "lot", scriptStr4, True)
    End Sub


    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim i As Integer

        If e.CommandName = "SplitItem" Then
            'If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                Dim rows_count As Integer = 0
                Dim rowNum As Integer = -1
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                If Not IsNothing(gvRow) Then
                    rowNum = gvRow.RowIndex
                End If

                Dim seq_string As String = "select MAX(convert(int,GRD_SEQ)) + 1 from WMS_GOODSRCV_D " & _
                                   "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                   "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                   "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

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

                dt.Rows(rowNum + 1).Item("grd_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rowNum + 1).Item("GRD_DISP_SEQ") = ViewState("n_cur_seq").ToString
                'dt.Rows(rowNum + 1).Item("GRD_DISP_SEQ") = dt.Rows(rowNum).Item("GRD_DISP_SEQ")

                dt.Rows(rowNum + 1).Item("GRD_ITM_CODE") = dt.Rows(rowNum).Item("GRD_ITM_CODE")
                dt.Rows(rowNum + 1).Item("ITM_SKU_NO") = dt.Rows(rowNum).Item("ITM_SKU_NO")
                dt.Rows(rowNum + 1).Item("GRD_PACK_KEY") = dt.Rows(rowNum).Item("GRD_PACK_KEY")
                dt.Rows(rowNum + 1).Item("GRD_PACK") = dt.Rows(rowNum).Item("GRD_PACK")
                dt.Rows(rowNum + 1).Item("GRD_ITM_NAME") = dt.Rows(rowNum).Item("GRD_ITM_NAME")
                'dt.Rows(rowNum + 1).Item("ITM_DESC") = dt.Rows(rowNum).Item("ITM_DESC")
                dt.Rows(rowNum + 1).Item("GRD_PO_QTY") = 0
                dt.Rows(rowNum + 1).Item("GRD_RCV_QTY") = 0
                dt.Rows(rowNum + 1).Item("GRD_OS_QTY") = 0
                'dt.Rows(rowNum + 1).Item("GRD_SERIES") = dt.Rows(rowNum).Item("GRD_SERIES")
                dt.Rows(rowNum + 1).Item("GRD_PACK_NO") = dt.Rows(rowNum).Item("GRD_PACK_NO")
                dt.Rows(rowNum + 1).Item("GRD_REF_NO") = dt.Rows(rowNum).Item("GRD_REF_NO")
                dt.Rows(rowNum + 1).Item("GRD_STOCKTYPE") = dt.Rows(rowNum).Item("GRD_STOCKTYPE")
                'dt.Rows(rowNum + 1).Item("GRD_BRAND") = dt.Rows(rowNum).Item("GRD_BRAND")
                dt.Rows(rowNum + 1).Item("GRD_REF_SEQ") = CInt(dt.Rows(rowNum).Item("GRD_SEQ").ToString)
                dt.Rows(rowNum + 1).Item("GRD_PALLET_NO") = gU.decodeNull(dt.Rows(rowNum).Item("GRD_PALLET_NO").ToString.Trim, "000")
                dt.Rows(rowNum + 1).Item("GRD_PCS_PER_UOM") = "1"
                dt.Rows(rowNum + 1).Item("GRD_UOM") = dt.Rows(rowNum).Item("GRD_UOM")
                dt.Rows(rowNum + 1).Item("GRD_TOT_PCS") = "0"
                dt.Rows(rowNum + 1).Item("GRD_VND_CODE") = dt.Rows(rowNum).Item("GRD_VND_CODE")
                dt.Rows(rowNum + 1).Item("GRD_EXPIRY_DATE") = dt.Rows(rowNum).Item("GRD_EXPIRY_DATE")
                dt.Rows(rowNum + 1).Item("GRD_MANU_DATE") = dt.Rows(rowNum).Item("GRD_MANU_DATE")
                dt.Rows(rowNum + 1).Item("GRD_REJ_QTY") = 0
                dt.Rows(rowNum + 1).Item("GRD_REJ_REASON") = ""
                dt.Rows(rowNum + 1).Item("GRD_INSP_REQ") = dt.Rows(rowNum).Item("GRD_INSP_REQ")
                dt.Rows(rowNum + 1).Item("GRD_STATUS") = dt.Rows(rowNum).Item("GRD_STATUS")

                dt.Rows(rowNum + 1).Item("GRD_UOM2") = dt.Rows(rowNum).Item("GRD_UOM2")
                dt.Rows(rowNum + 1).Item("GRD_QTY2") = dt.Rows(rowNum).Item("GRD_QTY2")

                'dt.Rows(rowNum + 1).Item("GRD_ON_BEHALF") = dt.Rows(rowNum).Item("GRD_ON_BEHALF")
                dt.Rows(rowNum + 1).Item("GRD_WH") = dt.Rows(rowNum).Item("GRD_WH")

                dt.Rows(rowNum + 1).Item("GRD_LENGTH") = dt.Rows(rowNum).Item("GRD_LENGTH")
                dt.Rows(rowNum + 1).Item("GRD_WIDTH") = dt.Rows(rowNum).Item("GRD_WIDTH")
                dt.Rows(rowNum + 1).Item("GRD_HEIGHT") = dt.Rows(rowNum).Item("GRD_HEIGHT")
                dt.Rows(rowNum + 1).Item("GRD_KG") = dt.Rows(rowNum).Item("GRD_KG")

                dt.Rows(rowNum + 1).Item("ITM_SERIAL_NO_YN") = dt.Rows(rowNum).Item("ITM_SERIAL_NO_YN")

                dt.Rows(rowNum + 1).Item("GRD_INSP_REQ") = dt.Rows(rowNum).Item("GRD_INSP_REQ")

                dt.Rows(rowNum + 1).Item("GRD_PCS_PER_CARTON") = dt.Rows(rowNum).Item("GRD_PCS_PER_CARTON")
                dt.Rows(rowNum + 1).Item("GRD_DESTINATION") = dt.Rows(rowNum).Item("GRD_DESTINATION")

                dt.Rows(rowNum + 1).Item("mFlag") = "N"

                dt.AcceptChanges()

                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
            ' End If

        ElseIf e.CommandName = "SHOW_SERIAL" Then

            Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

            DS_ITEM_CODE.Text = CType(gvRow.FindControl("grd_itm_code"), Label).Text
            DS_PACK_KEY.Text = CType(gvRow.FindControl("grd_pack_key"), Label).Text
            DS_SKU_NO.Text = CType(gvRow.FindControl("itm_sku_no"), Label).Text
            DS_ITEM_NAME.Text = CType(gvRow.FindControl("grd_itm_name"), Label).Text
            DS_BATCH_NO.Text = CType(gvRow.FindControl("grd_batch_no"), TextBox).Text
            DS_PALLET_NO.Text = CType(gvRow.FindControl("grd_pallet_no"), TextBox).Text
            DS_WH.Value = CType(gvRow.FindControl("grd_wh"), Label).Text

            sn_dt = Session("sn_dt")

            gvDetailSerialList.DataSource = sn_dt
            gvDetailSerialList.DataBind()

            updtPnl_DetailSerialItem.Update()
            updtPnl_DetailSerialList.Update()
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "SHOW_DETAIL_SERIAL", "$find('detailSerial_behavior').show();", True)



        End If
    End Sub

    Private Sub showSerial()
        'Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

        Dim gvRow As GridViewRow

        gvRow = GridView1.Rows(CInt(pbGvRowIndex.Value))

        DS_ITEM_CODE.Text = CType(gvRow.FindControl("grd_itm_code"), Label).Text
        DS_PACK_KEY.Text = CType(gvRow.FindControl("grd_pack_key"), Label).Text
        DS_SKU_NO.Text = CType(gvRow.FindControl("itm_sku_no"), Label).Text
        DS_ITEM_NAME.Text = CType(gvRow.FindControl("grd_itm_name"), Label).Text
        DS_BATCH_NO.Text = CType(gvRow.FindControl("grd_batch_no"), TextBox).Text
        DS_PALLET_NO.Text = CType(gvRow.FindControl("grd_pallet_no"), TextBox).Text
        DS_WH.Value = CType(gvRow.FindControl("grd_wh"), Label).Text

        sn_dt = Session("sn_dt")

        gvDetailSerialList.DataSource = sn_dt
        gvDetailSerialList.DataBind()

        updtPnl_DetailSerialItem.Update()
        updtPnl_DetailSerialList.Update()
        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "SHOW_DETAIL_SERIAL", "$find('detailSerial_behavior').show();", True)


    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "Seq No.", "編號")
        Call cU.newChangeGVLabel(GridView1, "Vendor Code", "供應商號碼")
        Call cU.newChangeGVLabel(GridView1, "Internal Item Code WMS", "貨物編號")
        Call cU.newChangeGVLabel(GridView1, "Item Code", "貨物編號")
        Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        Call cU.newChangeGVLabel(GridView1, "Pack Type", "封裝類型")
        Call cU.newChangeGVLabel(GridView1, "Pack No.", "封裝編號")
        Call cU.newChangeGVLabel(GridView1, "PO Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "Rcv Qty", "收貨數量")
        Call cU.newChangeGVLabel(GridView1, "OS Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "Pallet No.", "板號")
        Call cU.newChangeGVLabel(GridView1, "No. of Carton", "外箱數量")
        Call cU.newChangeGVLabel(GridView1, "Qty Per Carton", "外箱每件數量")
        Call cU.newChangeGVLabel(GridView1, "Carton No.", "外箱編號")
        Call cU.newChangeGVLabel(GridView1, "Lot", "批號")
        Call cU.newChangeGVLabel(GridView1, "Reference", "參考編號")
        'Call cU.newChangeGVLabel(GridView1, "Brand", "品牌")
        'Call cU.newChangeGVLabel(GridView1, "Series", "系列")
        'Call cU.newChangeGVLabel(GridView1, "Model", "模型")
        Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        Call cU.newChangeGVLabel(GridView1, "Size (LxWxH)", "長X寬X高")
        Call cU.newChangeGVLabel(GridView1, "Weight(kg)", "重量(kg)")
        Call cU.newChangeGVLabel(GridView1, "Sub-Total (kg)", "共(kg)")
        Call cU.newChangeGVLabel(GridView1, "Volume (cbm)", "體積(cbm)")
        Call cU.newChangeGVLabel(GridView1, "Destination", "目的地")
        Call cU.newChangeGVLabel(GridView1, "Stock Type", "貨品類型")
        Call cU.newChangeGVLabel(GridView1, "Number Per UOM", "單位件數")
        Call cU.newChangeGVLabel(GridView1, "Total Number", "總件數")
        Call cU.newChangeGVLabel(GridView1, "Expiry Date", "失效日期")
        Call cU.newChangeGVLabel(GridView1, "Manufactory Date", "生產日期")
        Call cU.newChangeGVLabel(GridView1, "Rejected Qty", "拒絕數量")
        Call cU.newChangeGVLabel(GridView1, "Rejected Reason", "拒絕理由")
        Call cU.newChangeGVLabel(GridView1, "Rej. Remark", "拒絕注明")
        Call cU.newChangeGVLabel(GridView1, "Insp. Req", "需要驗貨")
        Call cU.newChangeGVLabel(GridView1, "Status", "狀態")
        Call cU.newChangeGVLabel(GridView1, "UOM2", "單位2")
        Call cU.newChangeGVLabel(GridView1, "Qty2", "數量2")
        'Call cU.newChangeGVLabel(GridView1, "On Behalf", "代表")
        Call cU.newChangeGVLabel(GridView1, "Subinventory", "子庫存")
        'Call cU.newChangeGVLabel(GridView1, "", "")
        'Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        'Select Case e.Row.RowType
        '    Case DataControlRowType.Header
        'Dim oGridView As GridView = DirectCast(sender, GridView)
        'Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

        'REM **********************
        'REM Use for re-create the label to change the Langauge
        'REM Modify Here
        'Call cU.changeGVLabel(oGridViewRow, e, "", "")
        'Call cU.changeGVLabel(oGridViewRow, e, "", "")
        'Call cU.changeGVLabel(oGridViewRow, e, "", "")
        'Call cU.changeGVLabel(oGridViewRow, e, "Seq No.", "編號", HorizontalAlign.Center)
        'Call cU.changeGVLabel(oGridViewRow, e, "Vendor Code", "供應商號碼")
        'Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
        'Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
        'Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
        'Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名稱")
        'Call cU.changeGVLabel(oGridViewRow, e, "Pack Type", "封裝形式")
        'Call cU.changeGVLabel(oGridViewRow, e, "Pack No.", "封裝編號")
        'Call cU.changeGVLabel(oGridViewRow, e, "RO Qty", "訂單數量", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Received Qty", "收貨數量", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "OS Qty", "未付數量", HorizontalAlign.Right)
        ''Call cU.changeGVLabel(oGridViewRow, e, "WH", "倉庫")
        ''Call cU.changeGVLabel(oGridViewRow, e, "Loc", "位置")
        'Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號")
        'Call cU.changeGVLabel(oGridViewRow, e, "Carton Qty", "外箱數量", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Qty per Carton", "每箱件數", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "外箱編號")
        ''Call cU.changeGVLabel(oGridViewRow, e, "Lot No.", "批號")
        'Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
        ''Call cU.changeGVLabel(oGridViewRow, e, "Batch Date", "批號日期")
        'Call cU.changeGVLabel(oGridViewRow, e, "Reference", "文件編號")
        'Call cU.changeGVLabel(oGridViewRow, e, "Brand", "品牌")
        'Call cU.changeGVLabel(oGridViewRow, e, "Series", "系列")
        'Call cU.changeGVLabel(oGridViewRow, e, "Model", "型號")
        'Call cU.changeGVLabel(oGridViewRow, e, "UOM", "單位")
        'Call cU.changeGVLabel(oGridViewRow, e, "Size<br />(LxWxH in cm)", "大小<br />(LxWxH in cm)", HorizontalAlign.Center)
        'Call cU.changeGVLabel(oGridViewRow, e, "Weight(kg)", "重量(kg)", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Sub-Total kg", "重量(kg)", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Volume(cbm)", "體積(cbm)", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Destination", "目的地")
        'Call cU.changeGVLabel(oGridViewRow, e, "Stock Type", "存貨類型")
        'Call cU.changeGVLabel(oGridViewRow, e, "Number per UOM", "單位件數", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Total Number", "總件數", HorizontalAlign.Right)
        'Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "失效日期")
        'Call cU.changeGVLabel(oGridViewRow, e, "Manufactory Date", "生產日期")
        'Call cU.changeGVLabel(oGridViewRow, e, "Rejected Qty.", "退貨數量")
        'Call cU.changeGVLabel(oGridViewRow, e, "Status", "狀態")
        'Call cU.changeGVLabel(oGridViewRow, e, "", "")
        'Call cU.changeGVLabel(oGridViewRow, e, "", "")
        ''Call cU.changeGVLabel(oGridViewRow, e, "", "")
        'REM **********************

        'oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select

        'Dim gvColdt As DataTable
        'Dim paP As GlobalDBFunc.DBCmdPara
        'Dim selectSQL As String = ""

        'paP = New GlobalDBFunc.DBCmdPara
        'selectSQL = "Select upper(FLDO_FILED_CTRL) as FLDO_FILED_CTRL from WMS_FIELD_OPTION " & _
        '            "Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(ViewState("STORER_CODE")) & " AND FUN_CODE='IB_GR' AND ISNULL(FLDO_FIELD_OPTION, '') <> 'Y' AND FLDO_FIELD_TYPE is null " & _
        '            " AND FLDO_TABLE_NAME='WMS_GOODSRCV_D' "

        'gvColdt = gDB.getDataTable(selectSQL, , , , paP)

        'If gvColdt.Rows.Count > 0 Then
        '    Select Case e.Row.RowType
        '        Case DataControlRowType.Header
        '            For x = 0 To gvCol.Length - 1
        '                If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then
        '                    e.Row.Cells(x).Visible = False
        '                End If
        '            Next

        '        Case DataControlRowType.DataRow
        '            'For i = 0 To gvColdt.Rows.Count - 1
        '            '    ctrl = e.Row.FindControl(gvColdt.Rows(i).Item("FLDO_FILED_CTRL").ToString.Trim)

        '            '    If Not ctrl Is Nothing Then
        '            '        ctrl.Visible = False
        '            '    End If
        '            'Next
        '            Dim ctrlType As String = ""
        '            For x = 0 To gvCol.Length - 1
        '                If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then

        '                    If e.Row.Cells(x).Controls.Count > 0 Then
        '                        For Each ctl As Control In e.Row.Cells(x).Controls
        '                            'Set control invisible
        '                            Select Case TypeName(ctl).ToUpper
        '                                Case "TEXTBOX"
        '                                    DirectCast(ctl, TextBox).Enabled = False
        '                                Case "BUTTON"
        '                                    DirectCast(ctl, Button).Enabled = False

        '                                Case "IMAGEBUTTON"
        '                                    DirectCast(ctl, ImageButton).Enabled = False

        '                                Case "DROPDOWNLIST"
        '                                    DirectCast(ctl, DropDownList).Enabled = False

        '                                Case "COMBOBOX"
        '                                    DirectCast(ctl, AjaxControlToolkit.ComboBox).Enabled = False

        '                                Case "CHECKBOX"
        '                                    DirectCast(ctl, CheckBox).Enabled = False

        '                                Case "CHECKBOXLIST"
        '                                    For Z = 0 To DirectCast(ctl, CheckBoxList).Items.Count - 1
        '                                        DirectCast(ctl, CheckBoxList).Items(Z).Enabled = False
        '                                    Next

        '                                Case "IMAGE"
        '                                    DirectCast(ctl, Image).Visible = False
        '                            End Select
        '                        Next
        '                    End If

        '                    e.Row.Cells(x).Visible = False
        '                End If
        '            Next
        '        Case DataControlRowType.Footer
        '            For x = 0 To gvCol.Length - 1
        '                If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then
        '                    e.Row.Cells(x).Visible = False
        '                End If
        '            Next
        '    End Select
        'End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("grd_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_DISP_SEQ").ToString.Trim
                CType(e.Row.FindControl("grd_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRD_SEQ").ToString.Trim
                CType(e.Row.FindControl("grd_ref_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRD_REF_SEQ").ToString.Trim
                CType(e.Row.FindControl("grd_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("grd_itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRD_ITM_NAME").ToString.Trim
                CType(e.Row.FindControl("grd_vnd_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "grd_vnd_code").ToString.Trim
                CType(e.Row.FindControl("dsp_grd_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "grd_vnd_code").ToString.Trim

                uiFun.load_dropdownBy_ColCode(CType(e.Row.FindControl("grd_pack"), DropDownList), "WMS_GOODSRCV_D.GRD_PACK", Session("gLang"), , Session("gSelectLabel"))

                CType(e.Row.FindControl("grd_pack"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "GRD_PACK").ToString.Trim

                CType(e.Row.FindControl("grd_pack_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_PACK_NO").ToString.Trim
                CType(e.Row.FindControl("grd_po_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_PO_QTY").ToString.Trim)
                CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_RCV_QTY").ToString.Trim.Trim)
                CType(e.Row.FindControl("grd_os_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_OS_QTY").ToString.Trim)

                CType(e.Row.FindControl("grd_pallet_no"), TextBox).Text = gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "GRD_PALLET_NO").ToString.Trim, "000")
                CType(e.Row.FindControl("grd_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_CARTON_NO").ToString.Trim

                'CType(e.Row.FindControl("grd_batch_cd"), TextBox).Text = cU.chgToYYYYMMDD(DataBinder.Eval(e.Row.DataItem, "GRD_BATCH_CD").ToString.Trim)
                'CType(e.Row.FindControl("grd_lot_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_LOT_NO").ToString.Trim

                'uiFun.load_dropdown(CType(e.Row.FindControl("grd_batch_no"), DropDownList), "select dc_date_code from wms_date_code order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("grd_batch_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "grd_batch_no").ToString.Trim
                Dim ITM_STACKABLE_YN = "Y"
                Dim sqlQuery = "Select IsNull(ITM_STACKABLE_YN,'Y')ITM_STACKABLE_YN from WMS_ITEM Where IMP_CODE='" & Session("IMP_CODE") & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ITM_CODE='" & DataBinder.Eval(e.Row.DataItem, "GRD_ITM_CODE").ToString.Trim & "' and PACK_KEY='1'"
                Dim tmpdt = gDB.getDataTable(sqlQuery)
                If tmpdt IsNot Nothing AndAlso tmpdt.Rows.Count > 0 Then
                    ITM_STACKABLE_YN = tmpdt.Rows(0)("ITM_STACKABLE_YN")
                End If
                If ITM_STACKABLE_YN = "N" Then
                    Dim listItem As New ListItem
                    listItem.Value = "19000101"
                    listItem.Text = "19000101"
                    'CType(e.Row.FindControl("GRD_batch_no"), TextBox).Text = "19000101"
                    CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).Text = "31/12/2050"
                    CType(e.Row.FindControl("GRD_batch_no"), TextBox).Text = False
                    CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).Enabled = False
                Else
                    'CType(e.Row.FindControl("GRD_batch_no"), TextBox).Text = True
                    CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).Enabled = True
                    'uiFun.load_ComboBox(CType(e.Row.FindControl("GRD_BATCH_NO"), AjaxControlToolkit.ComboBox),
                    '                "select '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "GRd_batch_no").ToString.Trim) & "' as dc_date_code  union select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", , False)
                    CType(e.Row.FindControl("GRD_batch_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRd_batch_no").ToString.Trim
                    CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_EXPIRY_DATE").ToString.Trim
                End If


                'CType(e.Row.FindControl("grd_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRD_BATCH_NO").ToString.Trim

                CType(e.Row.FindControl("grd_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_REF_NO").ToString.Trim
                'CType(e.Row.FindControl("grd_brand"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_BRAND").ToString.Trim
                'CType(e.Row.FindControl("grd_series"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_SERIES").ToString.Trim
                'CType(e.Row.FindControl("grd_model"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_MODEL").ToString.Trim

                'Dim n1DropDown As DropDownList = CType(e.Row.FindControl("grd_uom"), DropDownList)
                'uiFun.load_dropdown(n1DropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'n1DropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "grd_uom").ToString.Trim
                'CType(e.Row.FindControl("grd_uom"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "grd_uom").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("grd_uom"), DropDownList), dtUOM, "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRD_UOM").ToString.Trim)

                uiFun.load_dropdown(CType(e.Row.FindControl("GRD_UOM2"), DropDownList), dtUOM2, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRD_UOM2").ToString.Trim)

                CType(e.Row.FindControl("grd_length"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_LENGTH").ToString.Trim)
                CType(e.Row.FindControl("grd_width"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_WIDTH").ToString.Trim)
                CType(e.Row.FindControl("grd_height"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_HEIGHT").ToString.Trim)
                CType(e.Row.FindControl("grd_kg"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_KG").ToString.Trim)

                CType(e.Row.FindControl("grd_destination"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_DESTINATION").ToString.Trim
                CType(e.Row.FindControl("grd_stocktype"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_STOCKTYPE").ToString.Trim

                'Dim nDropDown As DropDownList = CType(e.Row.FindControl("grd_cat"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "Select PO_CATEGORY_CODE,PO_CATEGORY_DESC from PO_CATEGORY_MASTER ORDER BY 2", "PO_CATEGORY_CODE", "PO_CATEGORY_DESC", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "grd_cat").ToString.Trim

                CType(e.Row.FindControl("grd_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim

                'CType(e.Row.FindControl("itm_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_DESC").ToString.Trim

                CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_PCS_PER_UOM").ToString.Trim
                CType(e.Row.FindControl("grd_tot_pcs"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_TOT_PCS").ToString.Trim



                'If ViewState("STO_BATCH_FIELD_REF") <> "" Then
                '    CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).Attributes.Add("onchange", "updtBatchNo('" & ViewState("STO_BATCH_FIELD_REF") & "', 'E', '" & CType(e.Row.FindControl("GRD_EXPIRY_DATE"), TextBox).ClientID & "', '" & CType(e.Row.FindControl("GRD_MANU_DATE"), TextBox).ClientID & "', '" & CType(e.Row.FindControl("GRD_batch_no"), AjaxControlToolkit.ComboBox).ClientID & "');")
                'End If

                CType(e.Row.FindControl("GRD_MANU_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_MANU_DATE").ToString.Trim

                CType(e.Row.FindControl("GRD_REJ_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_REJ_QTY").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("GRD_REJ_REASON"), DropDownList), dtRejReason, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRD_REJ_REASON").ToString.Trim, , True)

                CType(e.Row.FindControl("GRD_INSP_REQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRD_INSP_REQ").ToString.Trim
                If DataBinder.Eval(e.Row.DataItem, "GRD_INSP_REQ").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("dsp_GRD_INSP_REQ"), Label).Text = "Yes"
                Else
                    CType(e.Row.FindControl("dsp_GRD_INSP_REQ"), Label).Text = "No"
                End If

                'CType(e.Row.FindControl("GRD_STATUS"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "GRD_STATUS").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("GRD_STATUS"), DropDownList), dtDtlStatus, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRD_STATUS").ToString.Trim)

                'If DataBinder.Eval(e.Row.DataItem, "GRD_STATUS").ToString.Trim <> "" Then
                '    CType(e.Row.FindControl("dsp_GRD_STATUS"), Label).Text = CType(e.Row.FindControl("GRD_STATUS"), DropDownList).SelectedItem.Text
                'End If

                CType(e.Row.FindControl("GRD_QTY2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRD_QTY2").ToString.Trim

                'CType(e.Row.FindControl("GRD_ON_BEHALF"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRD_ON_BEHALF").ToString.Trim
                'If DataBinder.Eval(e.Row.DataItem, "GRD_ON_BEHALF").ToString.Trim = "Y" Then
                '    CType(e.Row.FindControl("dsp_GRD_ON_BEHALF"), Label).Text = "Yes"
                'ElseIf DataBinder.Eval(e.Row.DataItem, "GRD_ON_BEHALF").ToString.Trim = "T" Then
                '    CType(e.Row.FindControl("dsp_GRD_ON_BEHALF"), Label).Text = "IN-TRANS"
                'Else
                '    CType(e.Row.FindControl("dsp_GRD_ON_BEHALF"), Label).Text = "No"
                'End If

                CType(e.Row.FindControl("GRD_WH"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRD_WH").ToString.Trim
                'uiFun.load_dropdown(CType(e.Row.FindControl("GRD_WH"), DropDownList), dtWh, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRD_WH").ToString.Trim, , True)

                'CType(e.Row.FindControl("GRD_WH"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "GRD_WH").ToString.Trim

                CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "GRD_PCS_PER_CARTON").ToString.Trim)

                If DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("btnItemSerial"), ImageButton).Visible = True
                Else
                    CType(e.Row.FindControl("btnItemSerial"), ImageButton).Visible = False
                End If

                CType(e.Row.FindControl("btnItemSerial"), ImageButton).OnClientClick = "showSerial('" & e.Row.RowIndex & "');return false;"


                'Dim sImage As Image = CType(e.Row.FindControl("item_serial"), Image)
                'sImage.Attributes.Add("onmousedown", "MM_swapImage('" & gU.jsHTMLEncode(sImage.ClientID) & "','','../../images/btn_search_over.gif',1)")

                'If GR_STATUS.Text = "POSTED" Then
                '    sImage.Attributes.Add("onclick", "SerialLookUp(" & _
                '            "'" & DataBinder.Eval(e.Row.DataItem, "GRD_ITM_CODE").ToString.Trim & "', " & _
                '            "'" & DataBinder.Eval(e.Row.DataItem, "GRD_ITM_NAME").ToString.Trim & "', " & _
                '            "'" & DataBinder.Eval(e.Row.DataItem, "GRD_RCV_QTY").ToString.Trim & " ')")
                'Else
                '    sImage.Attributes.Add("onclick", "SerialLookUp(" & _
                '            "document.myform.document.getElementById(""" & CType(e.Row.FindControl("grd_itm_code"), Label).ClientID & """).innerHTML, " & _
                '            "document.myform.document.getElementById(""" & CType(e.Row.FindControl("grd_itm_name"), Label).ClientID & """).innerHTML, " & _
                '            "document.myform." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value)")

                'End If

                CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value * this.value;sumAll();")

                CType(e.Row.FindControl("grd_length"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                CType(e.Row.FindControl("grd_width"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                CType(e.Row.FindControl("grd_height"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                CType(e.Row.FindControl("grd_kg"), TextBox).Attributes("onkeypress") = "return maskKey(event)"


                CType(e.Row.FindControl("grd_length"), TextBox).Attributes.Add("onkeyup", _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                        "fixDecimal((this.value * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                        "sumAll();")


                CType(e.Row.FindControl("grd_width"), TextBox).Attributes.Add("onkeyup", _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                        "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                        "this.value * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                        "sumAll();")


                CType(e.Row.FindControl("grd_height"), TextBox).Attributes.Add("onkeyup", _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                        "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                        "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                        "this.value) / 1000000 * document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                        "sumAll();")

                'CType(e.Row.FindControl("grd_carton_no"), TextBox).Attributes.Add("onchange", "javascript:splitCarton(this.value);")

                If IsDBNull(DataBinder.Eval(e.Row.DataItem, "GRD_CBM")) Then
                    CType(e.Row.FindControl("grd_cbm"), TextBox).Text = "0"
                Else
                    CType(e.Row.FindControl("grd_cbm"), TextBox).Text = CDbl(DataBinder.Eval(e.Row.DataItem, "GRD_CBM")).ToString("###,###,##0.0000")
                End If

                'CType(e.Row.FindControl("grd_cbm"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "GRD_CBM").ToString.Trim)

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" AndAlso DataBinder.Eval(e.Row.DataItem, "GRD_NO_OF_CARTON").ToString.Trim = "" Then
                    CType(e.Row.FindControl("grd_no_of_carton"), TextBox).Text = 1
                    SumCar(1)
                Else
                    CType(e.Row.FindControl("grd_no_of_carton"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "GRD_NO_OF_CARTON").ToString.Trim)
                    SumCar(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_no_of_carton").ToString.Trim))
                End If

                SumPCS(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_tot_pcs").ToString.Trim))

                If cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_tot_pcs").ToString.Trim) > 0 Then
                    'SumKG(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_kg").ToString.Trim), cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_no_of_carton").ToString.Trim))
                    SumKG(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_kg").ToString.Trim), cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_rcv_qty").ToString.Trim))

                    If Not IsDBNull(DataBinder.Eval(e.Row.DataItem, "GRD_CBM")) Then
                        SumCBM(CDbl(DataBinder.Eval(e.Row.DataItem, "grd_cbm")))
                    End If
                    'SumCBM(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_cbm").ToString.Trim))
                End If

                CType(e.Row.FindControl("grd_tot_pcs"), TextBox).Attributes("onkeyup") = "sumAll();"

                'CType(e.Row.FindControl("grd_no_of_carton"), TextBox).Attributes("onkeyup") = _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                '                "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * this.value,3);" & _
                '                "calcKg('" & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"

                'CType(e.Row.FindControl("grd_no_of_carton"), TextBox).Attributes("onchange") = _
                '                "sumAll();"

                'CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).Attributes("onkeyup") = _
                '                "if (this.value!=0){document.forms[0]." & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & ".value = Math.ceil(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value / this.value);};" & _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                '                "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                '                "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * document.forms[0]." & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & ".value,3);" & _
                '                "calcKg('" & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"

                CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).Attributes("onkeyup") = _
                                "if (this.value!=0){document.forms[0]." & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & ".value = Math.ceil(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value / this.value);}"

                'CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).Attributes("onchange") = "sumAll();"

                'CType(e.Row.FindControl("grd_kg"), TextBox).Attributes("onkeyup") = "calcKg('" & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"
                CType(e.Row.FindControl("grd_kg"), TextBox).Attributes("onkeyup") = "calcKg('" & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"
                CType(e.Row.FindControl("grd_cbm"), TextBox).Attributes("onkeyup") = "sumAll();"

                REM **********************

                'Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                'If Session("gLang") = "E" Then
                '    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                '    'nButton.Text = "Delete"
                '    nButton.Text = "D"
                'ElseIf Session("gLang") = "C" Then
                '    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                '    'nButton.Text = "删除"
                '    nButton.Text = "删"
                'End If

                'Dim CopyButton As Button = CType(e.Row.FindControl("btnCopySize"), Button)

                'Dim jsCB As String = ""

                'jsCB += "document.forms[0]." & C_L.ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value;"
                'jsCB += "document.forms[0]." & C_W.ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value;"
                'jsCB += "document.forms[0]." & C_H.ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value;"
                'jsCB += "document.forms[0]." & C_KG.ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & ".value;"
                'jsCB += "return false;"

                'CopyButton.Attributes.Add("onclick", jsCB)

                'If Session("gLang") = "E" Then
                '    'CopyButton.Text = "Copy Size"
                '    CopyButton.Text = "C"
                'ElseIf Session("gLang") = "C" Then
                '    'CopyButton.Text = "複製大小"
                '    CopyButton.Text = "複"
                'End If

                'Dim PasteButton As Button = CType(e.Row.FindControl("btnPaste"), Button)

                'PasteButton.Attributes.Add("onclick", "javascript:copySizeTo(" & e.Row.RowIndex + 1 & "); document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                '       "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                '       "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                '        "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * " & _
                '        "document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                '        "calcKg('" & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();" & _
                '        "return false;")
                'PasteButton.Text = "Paste Size"

                If Session("gLang") = "E" Then
                    'PasteButton.Text = "P"
                ElseIf Session("gLang") = "C" Then
                    'PasteButton.Attributes.Add("onclick", "javascript:copySizeTo(" & e.Row.RowIndex + 1 & ");return false;")
                    'PasteButton.Text = "貼上大小"
                    'PasteButton.Text = "貼"
                End If

                Dim HighLightButton As Button = CType(e.Row.FindControl("btnHighlight"), Button)


                HighLightButton.Attributes.Add("onclick", "javascript:bindcolor(" & e.Row.RowIndex + 1 & ");return false;")

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "cb_select").ToString.Trim = "1" Then
                    CType(e.Row.FindControl("cb_select"), CheckBox).Checked = True
                End If

                If DataBinder.Eval(e.Row.DataItem, "GRD_REF_SEQ").ToString.Trim <> "" Then
                    DirectCast(e.Row.FindControl("btnSplit"), Button).Visible = False
                    DirectCast(e.Row.FindControl("grd_disp_seq"), TextBox).Visible = False

                    'CType(e.Row.FindControl("cb_select"), CheckBox).Visible = False

                    'DirectCast(e.Row.FindControl("grd_itm_code"), Label).Visible = False
                    'DirectCast(e.Row.FindControl("grd_itm_name"), Label).Visible = False

                    DirectCast(e.Row.FindControl("grd_po_qty"), TextBox).Visible = False
                    'DirectCast(e.Row.FindControl("grd_rcv_qty"), TextBox).Visible = False
                    DirectCast(e.Row.FindControl("grd_os_qty"), TextBox).Visible = False

                    If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then

                        For k As Integer = 0 To e.Row.Cells.Count - 1
                            e.Row.Cells(k).CssClass = "GV_SPLIT"
                        Next
                    End If

                    'CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).ClientID & ".value * this.value;sumAll();")

                    If editMode.Value <> "V" Then
                        Dim objName As String = "TM" & DataBinder.Eval(e.Row.DataItem, "GRD_REF_SEQ").ToString.Trim

                        CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes("onkeyup") = _
                                "if(document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).ClientID & ".value != 0){document.forms[0]." & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & ".value = Math.ceil(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value / document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).ClientID & ".value);}; " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                                "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                                "calcKg('" & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"


                        If firstROClientId <> "" AndAlso firstOSClientId <> "" Then
                            CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).ClientID & ".value * this.value;" & _
                                                                                                        "sumAll();")
                        Else
                            CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).ClientID & ".value * this.value;" & _
                                                                                                        "sumAll();")
                        End If

                        Dim jsString As String = objName & ".push(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ");"
                        Me.ClientScript.RegisterStartupScript(Me.GetType, "addAnVar_" & objName & "_" & DataBinder.Eval(e.Row.DataItem, "GRD_SEQ").ToString.Trim, jsString, True)
                    End If
                Else
                    'CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_os_qty"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_po_qty"), TextBox).ClientID & ".value - this.value;document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).ClientID & ".value * this.value;sumAll();")

                    If editMode.Value <> "V" Then
                        Dim objName As String = "TM" & DataBinder.Eval(e.Row.DataItem, "GRD_SEQ").ToString.Trim

                        CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes("onkeyup") = _
                                "if(document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).ClientID & ".value != 0){document.forms[0]." & CType(e.Row.FindControl("grd_no_of_carton"), TextBox).ClientID & ".value = Math.ceil(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value / document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_carton"), TextBox).ClientID & ".value);}; " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_cbm"), TextBox).ClientID & ".value = " & _
                                "fixDecimal((document.forms[0]." & CType(e.Row.FindControl("grd_length"), TextBox).ClientID & ".value * " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_width"), TextBox).ClientID & ".value * " & _
                                "document.forms[0]." & CType(e.Row.FindControl("grd_height"), TextBox).ClientID & ".value) / 1000000 * document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ".value,3);" & _
                                "calcKg('" & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & "','" & CType(e.Row.FindControl("sub_grd_kg"), Label).ClientID & "','" & CType(e.Row.FindControl("grd_kg"), TextBox).ClientID & "');sumAll();"


                        CType(e.Row.FindControl("grd_rcv_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("grd_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("grd_pcs_per_uom"), TextBox).ClientID & ".value * this.value;" & _
                                                                                "sumAll();")

                        firstROClientId = "document.forms[0]." & CType(e.Row.FindControl("grd_po_qty"), TextBox).ClientID
                        firstOSClientId = "document.forms[0]." & CType(e.Row.FindControl("grd_os_qty"), TextBox).ClientID

                        Dim jsString As String = "var " & objName & " = [];" & _
                                                    objName & ".push(document.forms[0]." & CType(e.Row.FindControl("grd_rcv_qty"), TextBox).ClientID & ");"

                        Me.ClientScript.RegisterStartupScript(Me.GetType, "addVar_" & objName & "_" & DataBinder.Eval(e.Row.DataItem, "GRD_SEQ").ToString.Trim, jsString, True)
                    End If
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim <> "N" AndAlso DataBinder.Eval(e.Row.DataItem, "GRD_REF_SEQ").ToString.Trim = "" Then
                    Dim hasValue As String = DB.getValueFromSQL("SELECT 1 FROM WMS_GOODSRCV_D " & _
                                                                "WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                                                "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                                "AND GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                                                "AND GRD_ITM_CODE = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "GRD_ITM_CODE").ToString.Trim) & "' " & _
                                                                "AND GRD_PACK_KEY = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "GRD_PACK_KEY").ToString.Trim) & "' " & _
                                                                "AND GRD_REF_SEQ is not null")

                    'If hasValue <> "" Then
                    '    nButton.Attributes("onclick") = "javascript:alert('This item has been splitted!\nYou must delete all splitted items before you can delete the original item.');return false;"
                    'End If
                End If

                ar.hideGVForStorer(GridView1, STORER_CODE.Text, "GR", "WMS_GOODSRCV_D", e)

                CType(e.Row.FindControl("sub_grd_kg"), Label).Text = FormatNumber(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_rcv_qty").ToString.Trim) * cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "grd_kg").ToString.Trim), 2)
            Case DataControlRowType.Footer
                CType(e.Row.FindControl("grd_sub_total_pcs"), Label).Text = GetTotal()
                CType(e.Row.FindControl("grd_total_carton"), Label).Text = GetCarTotal()
                CType(e.Row.FindControl("grd_total_kg"), Label).Text = FormatNumber(GetKGTotal(), 2)
                CType(e.Row.FindControl("grd_total_cbm"), Label).Text = FormatNumber(GetCBMTotal(), 2)
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int,GRD_SEQ)) + 1 from WMS_GOODSRCV_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("grd_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            Session("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Private Function validateAll(Optional ByRef sFlag As String = "") As Boolean
        Dim selectSql As String = ""
        Dim i, j As Integer

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If GR_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_GR_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", lbl_GR_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(GR_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid date, " & lbl_GR_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的日期, " & lbl_GR_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If PO_DATE.Text.Trim <> "" And Not gU.isValidDate(PO_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid date, " & lbl_PO_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的日期, " & lbl_PO_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(GR_TOT_PALLET.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, " & lbl_GR_TOT_PALLET.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, " & lbl_GR_TOT_PALLET.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If GR_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Too many characters, maximum length of " & lbl_GR_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "字數太多, " & lbl_GR_REM.Text & "最多只限200字!", Session("gLang"))
            End If
            Return False
        End If

        'If GR_DOC_NO.Value = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsgNew(updtPnlAlert, "", "Please Select PO before save!", Session("gLang"))
        '    Else
        '        uiFun.displayMsgNew(updtPnlAlert, "", "Please Select PO before save!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                'Dim loc As String = CType(GridView1.Rows(i).FindControl("grd_loc"), HiddenField).Value

                'If loc = "" Then
                '    If Session("gLang") = "E" Then
                '        uiFun.displayMsgNew(updtPnlAlert, "", "Location Cannot Be Empty!", Session("gLang"))
                '    Else
                '        uiFun.displayMsgNew(updtPnlAlert, "", "位置不能空白!", Session("gLang"))
                '    End If
                '    Return False
                'End If
                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_po_qty"), TextBox).Text) Then

                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, PO Qty!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 訂單數量!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Visible = True AndAlso _
                    Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text) Then

                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Received Qty!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 收貨數量!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("grd_os_qty"), TextBox).Visible = True AndAlso _
                    Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_os_qty"), TextBox).Text) Then

                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, OS Qty!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 未付數量!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_length"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Length!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小(L)!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_width"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Width!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小(W)!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_height"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Height!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小(H)!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_kg"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Weight(kg)!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 重量(kg)!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_cbm"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Volume(cbm)!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 體積(cbm)!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_pcs_per_uom"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Number per UOM!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 單位件數!", Session("gLang"))
                    End If
                    Return False
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_tot_pcs"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Total Number!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 總件數!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text.Trim <> "" Then
                    Dim strBatch As String = CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text
                    Dim batchDate As String = ""
                    If strBatch IsNot Nothing AndAlso strBatch.Length >= 8 Then
                        batchDate = strBatch.Substring(0, 4) + "-" + strBatch.Substring(4, 2) + "-" + strBatch.Substring(6, 2)
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!!", Session("gLang"))
                        End If
                        Return False
                    End If
                    Dim fromDateValue As DateTime
                    If DateTime.TryParse(batchDate, fromDateValue) Then

                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!!", Session("gLang"))
                        End If
                        Return False
                    End If
                End If

                'grd_no_of_carton
                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_no_of_carton"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, No. of Carton!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 總箱數!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("grd_rej_qty"), TextBox).Text.Trim <> "" AndAlso Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_rej_qty"), TextBox).Text.Trim) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Rejected Qty.!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 退貨數量!", Session("gLang"))
                    End If
                    Return False
                End If

                'If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_po_qty"), TextBox).Text, 0) < gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text, 0) Then
                '    If Session("gLang") = "E" Then
                '        uiFun.displayMsgNew(updtPnlAlert, "", " Received Qty cannot greater than PO Qty!", Session("gLang"))
                '    Else
                '        uiFun.displayMsgNew(updtPnlAlert, "", "收貨數量不能大於貨單數量!", Session("gLang"))
                '    End If
                '    Return False
                'End If

                If CType(GridView1.Rows(i).FindControl("grd_rej_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("grd_rej_qty"), TextBox).Text.Trim) <> 0 Then

                    Dim totalRejQty, totalRcvQty As Double


                    For j = 0 To GridView1.Rows.Count - 1
                        If CType(GridView1.Rows(i).FindControl("grd_ref_seq"), HiddenField).Value = CType(GridView1.Rows(j).FindControl("grd_seq"), HiddenField).Value Then

                            If CType(GridView1.Rows(j).FindControl("grd_po_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("grd_rej_qty"), TextBox).Text.Trim) + gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text.Trim, 0) > CDbl(CType(GridView1.Rows(j).FindControl("grd_po_qty"), TextBox).Text.Trim) Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsgNew(updtPnlAlert, "", "Rejected Qty cannot greater than PO Qty!", Session("gLang"))
                                Else
                                    uiFun.displayMsgNew(updtPnlAlert, "", "拒收數量不能大於貨單數量!", Session("gLang"))
                                End If
                                Return False
                            End If

                            Exit For
                        End If
                    Next

                    If CType(GridView1.Rows(i).FindControl("GRD_REJ_REASON"), DropDownList).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Please select Reject Reason!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "請選擇拒收原因!", Session("gLang"))
                        End If
                        Return False
                    End If
                End If

                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("grd_qty2"), TextBox).Text) Then

                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Qty2!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 數量2!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim) Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Invalid Manufactory Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Focus()
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim) Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Invalid Expiry Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Focus()
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim <> "" AndAlso CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim <> "" AndAlso _
                    Not gU.isEndDateLaterThanStart(CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim, CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim, True) Then

                    uiFun.displayMsgNew(updtPnlAlert, "", "Expiry Date must be later than Manufactory Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Focus()
                    Return False
                End If

                If sFlag = "P" Then

                End If
            Next
        End If

        If Not isValidItems() Then
            Return False
        End If

        updtBatchNo()

        updtPutAwayPallet()

        If Not isValidPutAway(sFlag) Then
            Return False
        End If

        If Not isValidItemsSerial(sFlag) Then
            Return False
        End If

        Return True

    End Function

    Private Function isValidItemsSerial(ByVal sFlag As String) As Boolean
        Dim snDt, pa_dt As DataTable
        Dim i, j As Integer
        Dim grsCodestr As String = ""
        Dim isValidLoc As Boolean

        If sFlag = "P" Then
            snDt = Session("sn_dt")

            pa_dt = Session("_M_IB_GR_TMP_pa_dt")

            For i = 0 To snDt.Rows.Count - 1
                If snDt(i).Item("GRS_SERIAL_NO").ToString.Trim = "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Please enter serail no. for item code:" & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "Please enter serail no. for item code:" & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    End If
                    Return False
                End If


                If snDt(i).Item("GRS_LOC").ToString.Trim = "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Item Serial location cannot be empty! Please check item serial of " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "Item Serial location cannot be empty! Please check item serial of " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    End If
                    Return False
                End If

                isValidLoc = False

                If snDt(i).Item("GRS_LOC").ToString.Trim = dmgLocValue Then
                    isValidLoc = True
                Else
                    For j = 0 To pa_dt.Rows.Count - 1
                        If pa_dt.Rows(j).Item("mFlag").ToString <> "D" Then
                            If snDt(i).Item("GRS_LOC").ToString.Trim = pa_dt.Rows(j).Item("GRA_LOC").ToString.Trim Then
                                isValidLoc = True
                                Exit For
                            End If
                        End If
                    Next
                End If

                If Not isValidLoc Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid Item Serial location! Please check item serial of " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "Invalid Item Serial location! Please check item serial of " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                    End If
                End If

                grsCodestr = gU.appendToList(grsCodestr, snDt.Rows(i).Item("GRS_ITM_CODE").ToString.Trim & "||" & snDt.Rows(i).Item("GRS_PACK_KEY").ToString.Trim)

            Next


            Dim tempDT As DataTable = Session("dt")

            If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    If tempDT.Rows(i).Item("ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                        If Not gU.inList(grsCodestr, tempDT.Rows(i).Item("GRD_ITM_CODE").ToString.Trim & "||" & tempDT.Rows(i).Item("GRD_PACK_KEY").ToString.Trim) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Please enter serial number for item code: " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "Please enter serial number for item code: " & snDt(i).Item("GRS_ITM_CODE").ToString.Trim & ".", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If
                Next
            End If

        End If

        Return True
    End Function

    Private Function isValidPutAway(Optional ByRef sFlag As String = "") As Boolean
        Dim itmDict, paItmDict As Dictionary(Of String, Double)
        Dim itmDictKey, paItmDictKey As Dictionary(Of String, Double).KeyCollection
        Dim keyArray As String()
        Dim itmKey As String
        Dim pa_dt As DataTable
        Dim i As Integer
        Dim javaStr As String
        Dim warnItmList1, warnItmList2, warnItmList3 As String
        Dim msg As String

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()

        If sFlag = "C" Then
            Return True
        End If

        javaStr = ""

        pa_dt = Session("_M_IB_GR_TMP_pa_dt")

        If sFlag = "RELEASE" AndAlso (pa_dt Is Nothing OrElse pa_dt.Rows.Count = 0) Then
            uiFun.displayMsgNew(updtPnlAlert, "", "Please generate Put Away first!", Session("gLang"))
            Return False
        End If


        If sFlag = "P" Then
            For i = 0 To pa_dt.Rows.Count - 1
                If pa_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                    If pa_dt.Rows(i).Item("GRA_PA_QTY") IsNot DBNull.Value And pa_dt.Rows(i).Item("GRA_REJ_QTY") IsNot DBNull.Value Then

                        If gU.decodeEmptyCdbl(pa_dt.Rows(i).Item("GRA_PA_QTY"), 0) + gU.decodeEmptyCdbl(pa_dt.Rows(i).Item("GRA_REJ_QTY"), 0) <= 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Put Away Qty must be greater than ZERO! Please check put away list.", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "上架數量要大過零! 請檢查儲存清單.", Session("gLang"))
                            End If
                            Return False
                        ElseIf gU.decodeEmptyCdbl(pa_dt.Rows(i).Item("GRA_PA_QTY"), 0) > 0 AndAlso pa_dt.Rows(i).Item("GRA_LOC").ToString.Trim = "" Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Put away location cannot be empty! Please check put away list.", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "儲存位置不能空白! 請檢查儲存清單.", Session("gLang"))
                            End If
                            Return False
                            'ElseIf pa_dt.Rows(i).Item("GRA_BATCH_NO").ToString.Trim = "" Then
                            '    If Session("gLang") = "E" Then
                            '        uiFun.displayMsgNew(updtPnlAlert, "", "Put away Batch No. cannot be empty! Please check put away list.", Session("gLang"))
                            '    Else
                            '        uiFun.displayMsgNew(updtPnlAlert, "", "Batch No.不能空白! 請檢查儲存清單.", Session("gLang"))
                            '    End If
                            '    Return False
                        End If
                    End If
                End If


            Next
        End If

        'Generate item dict for put away list
        If GridView1.Rows.Count > 0 Then
            itmDict = New Dictionary(Of String, Double)

            For i = 0 To GridView1.Rows.Count - 1
                'If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text, "0") > 0 AndAlso CType(GridView1.Rows(i).FindControl("GRD_ON_BEHALF"), HiddenField).Value <> "Y" Then
                'itmKey = CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text.Trim & "#_#" & _
                '            CType(GridView1.Rows(i).FindControl("grd_pack_key"), TextBox).Text.Trim & "#_#" & _
                '            CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim

                If CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Enabled = True Then
                    itmKey = CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text.Trim & "#_#" &
                                CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text.Trim & "#_#" &
                                CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim & "#_#" &
                                CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text

                    If Not itmDict.ContainsKey(itmKey) Then
                        itmDict.Add(itmKey, gU.decodeEmptyCdbl((CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text), 0))
                    Else
                        itmDict.Item(itmKey) = itmDict.Item(itmKey) + gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text, 0)
                    End If
                End If
            Next
        End If


        If itmDict.Count > 0 Then
                'warnItmList1 = ""
                'warnItmList2 = ""
                paItmDict = New Dictionary(Of String, Double)

                For i = 0 To pa_dt.Rows.Count - 1
                    'If gU.decodeEmptyCdbl(pa_dt.Rows(i).Item("GRA_PA_QTY"), 0) > 0 And pa_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                    If pa_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                        itmKey = pa_dt.Rows(i).Item("GRA_ITM_CODE").ToString & "#_#" & _
                                pa_dt.Rows(i).Item("GRA_PACK_KEY").ToString & "#_#" & _
                                pa_dt.Rows(i).Item("GRA_PALLET_NO").ToString & "#_#" & _
                                pa_dt.Rows(i).Item("GRA_BATCH_NO").ToString

                    Dim rej_Qty = If(pa_dt.Rows(i).Item("GRA_REJ_QTY") Is DBNull.Value, "0", pa_dt.Rows(i).Item("GRA_REJ_QTY"))
                    If Not paItmDict.ContainsKey(itmKey) Then
                        paItmDict.Add(itmKey, pa_dt.Rows(i).Item("GRA_PA_QTY") + CDbl(gU.decodeNullOrEmpty(rej_Qty.ToString.Trim, "0")))
                    Else
                        paItmDict.Item(itmKey) = paItmDict.Item(itmKey) + pa_dt.Rows(i).Item("GRA_PA_QTY") + gU.decodeEmptyCdbl(rej_Qty, 0)
                    End If

                        'If itmDict.ContainsKey(itmKey) Then
                        '    If paItmDict.Item(itmKey) > itmDict.Item(itmKey) Then
                        '        If sFlag = "P" OrElse (sFlag = "" AndAlso Request("moduleAction") <> "SAVEOK") Then
                        '            warnItmList1 = gU.appendToList(warnItmList1, pa_dt.Rows(i).Item("GRA_ITM_CODE").ToString)
                        '        End If
                        '    End If
                        'Else
                        '    If sFlag = "P" OrElse (sFlag = "" AndAlso Request("moduleAction") <> "SAVEOK") Then
                        '        warnItmList2 = gU.appendToList(warnItmList2, pa_dt.Rows(i).Item("GRA_ITM_CODE").ToString)
                        '    End If
                        'End If
                    End If
                Next


                warnItmList1 = ""
                warnItmList2 = ""

                paItmDictKey = paItmDict.Keys

                For i = 0 To paItmDictKey.Count - 1
                    'Check if put away qty <> receive qty
                    If itmDict.ContainsKey(paItmDictKey(i)) Then
                        If paItmDict.Item(paItmDictKey(i)) <> itmDict.Item(paItmDictKey(i)) Then
                            'If sFlag = "P" OrElse (sFlag = "" AndAlso Request("moduleAction") <> "SAVEOK") Then
                            If sFlag = "P" Then
                                keyArray = Split(paItmDictKey(i), "#_#")

                                msg = keyArray(0)

                                If keyArray.Length > 3 Then
                                    If keyArray(3) <> "" Then
                                        msg += "-" & keyArray(3)
                                    End If
                                End If

                                warnItmList1 = gU.appendToList(warnItmList1, msg)
                            End If
                        End If
                    Else
                        'Check if put away item not in receive item
                        If sFlag = "P" OrElse (sFlag = "" AndAlso moduleAction.Value <> "SAVEOK") Then
                            keyArray = Split(paItmDictKey(i), "#_#")

                            msg = keyArray(0)

                            If keyArray.Length > 3 Then
                                If keyArray(3) <> "" Then
                                    msg += "-" & keyArray(3)
                                End If
                            End If

                            warnItmList2 = gU.appendToList(warnItmList2, msg)
                        End If

                    End If
                Next

                warnItmList3 = ""

                itmDictKey = itmDict.Keys

                For i = 0 To itmDictKey.Count - 1
                    If Not paItmDict.ContainsKey(itmDictKey(i)) Then
                        keyArray = Split(itmDictKey(i), "#_#")

                        msg = keyArray(0)

                        If keyArray.Length > 3 Then
                            If keyArray(3) <> "" Then
                                msg += "-" & keyArray(3)
                            End If
                        End If

                        warnItmList3 = gU.appendToList(warnItmList3, msg)
                    End If
                Next

                If warnItmList3 <> "" Then
                    If sFlag = "P" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Put away list has not yet been assigned for Received item [" & warnItmList3 & "] . ", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "領取物件[" & warnItmList3 & "]不存在於儲存物件.", Session("gLang"))
                        End If
                        Return False
                    ElseIf sFlag = "" And moduleAction.Value <> "SAVEOK" Then
                        If javaStr <> "" Then
                            javaStr = javaStr & " && "
                        End If
                        If Session("gLang") = "E" Then
                            javaStr = javaStr & "confirm(""Put away list has not yet been assigned for Received item [" & warnItmList3 & "], confirm to proceed?"")"
                        Else
                            javaStr = javaStr & "confirm(""領取物件[" & warnItmList3 & "]不存在於儲存物件, 確定輸入?"")"
                        End If
                    End If
                End If

                If warnItmList2 <> "" Then
                    If sFlag = "P" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Put away item [" & warnItmList2 & "] is not found in receive items. ", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "儲存物件[" & warnItmList2 & "]不存在於領取物件.", Session("gLang"))
                        End If
                        Return False
                    ElseIf sFlag = "" And moduleAction.Value <> "SAVEOK" Then
                        If javaStr <> "" Then
                            javaStr = javaStr & " && "
                        End If
                        If Session("gLang") = "E" Then
                            javaStr = javaStr & "confirm(""Put away item [" & warnItmList2 & "] is not found in receive items, confirm to proceed?"")"
                        Else
                            javaStr = javaStr & "confirm(""儲存物件[" & warnItmList2 & "]不存在於領取物件, 確定輸入?"")"
                        End If
                    End If
                End If

                If warnItmList1 <> "" Then
                    If sFlag = "P" Then
                        If Session("gLang") = "E" Then
                            'uiFun.displayMsgNew(updtPnlAlert, "", "Item: [" & warnItmList1 & "] Put away qty is greater than received qty! Please check put away list.", Session("gLang"))
                            uiFun.displayMsgNew(updtPnlAlert, "", "Item: [" & warnItmList1 & "] Put away qty is different from received qty! Please check put away list.", Session("gLang"))
                        Else
                            'uiFun.displayMsgNew(updtPnlAlert, "", "物件: [" & warnItmList1 & "] 儲存數量高於領取數量! 請檢查儲存清單.", Session("gLang"))
                            uiFun.displayMsgNew(updtPnlAlert, "", "物件: [" & warnItmList1 & "] 儲存數量不同於領取數量! 請檢查儲存清單.", Session("gLang"))
                        End If
                        Return False
                    ElseIf sFlag = "" And moduleAction.Value <> "SAVEOK" Then
                        If javaStr <> "" Then
                            javaStr = javaStr & " && "
                        End If
                        If Session("gLang") = "E" Then
                            'javaStr = javaStr & "confirm(""Item: [" & warnItmList1 & "] Put away qty is greater than received qty, confirm to proceed?"")"
                            javaStr = javaStr & "confirm(""Item: [" & warnItmList1 & "] Put away qty is different from received qty, confirm to proceed?"")"
                        Else
                            'javaStr = javaStr & "confirm(""物件: [" & warnItmList1 & "] 儲存數量高於領取數量, 確定輸入?"")"
                            javaStr = javaStr & "confirm(""物件: [" & warnItmList1 & "] 儲存數量不同於領取數量, 確定輸入?"")"
                        End If
                    End If
                End If

            End If
        'End If

        If javaStr <> "" Then
            javaStr = "if(" & javaStr & ") saveok();"
            'If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
            '    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
            'End If
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "CONFIRM_WARN", javaStr, True)

            Return False
        End If

        Return True
    End Function

    Private Function isValidItems() As Boolean
        Dim i As Integer

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()

        For i = 0 To GridView1.Rows.Count - 1

            Dim query = From r In dt.Rows
                        Where r.item("grd_itm_code") = CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text _
                            And r.item("mFlag") <> "D"
                        Select r
            'And r.item("grd_pack_key") = CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text _
            'And r.item("grd_pack_key") = CType(GridView1.Rows(i).FindControl("grd_pack_key"), TextBox).Text _

            If query.Count > 0 Then
                If Not DB.chkItmExist(IMP_CODE.Value.Trim, STORER_CODE.SelectedValue, CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text, CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Item code is not existed in item master : " & CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text & "-" & CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text & "!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "物件號碼不存在系統中 : " & CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text & "-" & CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text & "!", Session("gLang"))
                    End If
                    Return False
                End If
            End If
        Next

        Return True
    End Function

    Protected Function save(Optional ByRef sFlag As String = "") As Boolean
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        Dim pa_dt As DataTable
        Dim dupSQL As String
        Dim dupTbl As New DataTable
        Dim updateSqL, insertSql As String
        Dim grd_batch_no As String
        Dim serial_seq As Integer
        Dim selectSql As String
        Dim tmpDt As DataTable
        Dim newDocNo, newTrackNo As String
        Dim updtRow As Integer
        Dim i As Integer

        If validateAll(sFlag) Then
            pa_dt = Session("_M_IB_GR_TMP_pa_dt")

            sn_dt = Session("sn_dt")

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("GR", gConn, transaction)
                    REM **********************

                    sql_string = "select count(*) from WMS_GOODSRCV " & _
                                "where IMP_CODE = '" & Session("imp_code") & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "'"

                    If CInt(DB.getValueFromSQL(sql_string)) > 0 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Duplicate key! " & lbl_GR_CODE.Text & "(" & GR_CODE.Text.Trim & ") is already existed!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "資料重複! " & lbl_GR_CODE.Text & "(" & GR_CODE.Text.Trim & ")已经存在!", Session("gLang"))
                        End If
                        transaction.Rollback()
                        Return False
                    End If

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "insert into WMS_GOODSRCV " &
                    "(IMP_CODE, STORER_CODE, GR_CODE, PRJ_CODE, GR_STATUS, GR_DATE, GR_RCV_BY, GR_DOC_NO, " &
                     "GR_DOC_TYPE, PO_DATE, VND_CODE, VND_NAME, GR_VND_DNREF, GR_REM, GR_DESTINATION, GR_TOT_PALLET, GR_TRACK_NO, GR_REF_NO, GR_HAWB, GR_INV_NO,GR_TRANS_TYPE,GR_CAT," &
                     "GR_EDI_PO_NO, GR_WH_CODE, " &
                     "sys_cb, sys_cd, sys_lub, sys_lud)" &
                    "values ( " &
                    "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                    "'" & gU.dbEncode(PRJ_CODE.SelectedValue) & "', '" & gU.dbEncode(GR_STATUS.Text.Trim) & "', " & gU.convdbDate(gU.dbEncode(GR_DATE.Text.Trim)) & ", " &
                    "'" & gU.dbEncode(GR_RCV_BY.Text.Trim) & "', '" & gU.dbEncode(GR_DOC_NO.Value.Trim) & "', '" & gU.dbEncode(GR_DOC_TYPE.Value.Trim) & "', " &
                    gU.convdbDate(gU.dbEncode(PO_DATE.Text.Trim)) & ", '" & gU.dbEncode(dsp_VND_CODE.Text.Trim) & "', N'" & gU.dbEncode(VND_NAME.Value) & "', " &
                    "'" & gU.dbEncode(GR_VND_DNREF.Text.Trim) & "', N'" & gU.dbEncode(GR_REM.Text.Trim) & "', " &
                    "N'" & gU.dbEncode(GR_DESTINATION.Text.Trim) & "', '" & gU.dbEncode(gU.decodeNullOrEmpty(GR_TOT_PALLET.Text.Trim, "0")) & "', " &
                    "'" & gU.dbEncode(GR_TRACK_NO.Value) & "', '" & gU.dbEncode(GR_REF_NO.Value) & "', '" & gU.dbEncode(GR_HAWB.Text.Trim) & "', '" & gU.dbEncode(GR_INV_NO.Text.Trim) & "', '" & gU.dbEncode(GR_TRANS_TYPE.SelectedValue) & "','" & gU.dbEncode(gr_cat.SelectedValue) & "'," &
                    "'" & gU.dbEncode(GR_EDI_PO_NO.Text.Trim) & "', '" & gU.dbEncode(GR_WH_CODE.Text.Trim) & "', " &
                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '"N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.SelectedValue) & "', N'" & nextNo & "', " & _
                    REM **********************

                    REM Generate Document Link 
                    dl.genDocLink("GR", nextNo, "RO", GR_DOC_NO.Value, STORER_CODE.SelectedValue, gConn, transaction)

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        reOrderGRD(dt)
                        uiFun.reOrderDetails(dt, "grd_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""

                            grd_batch_no = ""
                            grd_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("grd_batch_no").ToString.Trim, gU.decodeNull(rows.Item("GRD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("GRD_EXPIRY_DATE").ToString.Trim, ""))

                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    If grd_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(grd_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSqL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                                        "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                        "'" & gU.dbEncode(grd_batch_no) & "'," & _
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSqL, gConn, transaction)
                                        End If
                                    End If


                                    itemSQL = "insert into WMS_GOODSRCV_D " &
                                            "(imp_code, storer_code, gr_code, grd_seq, grd_disp_seq, grd_pack, grd_pack_no, grd_pallet_no, " &
                                             "grd_carton_no, grd_batch_no, grd_ref_no, grd_itm_code, grd_brand, grd_series, " &
                                             "grd_model, grd_itm_name, grd_po_qty, grd_rcv_qty, grd_os_qty, " &
                                             "grd_uom, grd_width, grd_length, grd_height, grd_kg, grd_cbm, " &
                                             "grd_destination, grd_stocktype, grd_pack_key, grd_pcs_per_uom, grd_tot_pcs, grd_ref_seq, " &
                                             "grd_no_of_carton, grd_pcs_per_carton, grd_vnd_code, grd_expiry_date, grd_manu_date, " &
                                             "grd_rej_qty, grd_rej_reason, grd_insp_req, " &
                                             "grd_status, grd_on_behalf, grd_wh, grd_uom2, grd_qty2, GRD_REJ_RMKS, " &
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                            "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_disp_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pallet_no").ToString.Trim, "000")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_carton_no").ToString.Trim, "")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(grd_batch_no)) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_ref_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_brand").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_series").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_model").ToString.Trim, "")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_name").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_po_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rcv_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_os_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_uom").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_width").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_length").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_height").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_kg").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_cbm").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_destination").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_stocktype").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_uom").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_tot_pcs").ToString.Trim, "0")) & "', " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_ref_seq").ToString.Trim, "null")) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_no_of_carton").ToString.Trim, "0")) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_carton").ToString.Trim, "0")) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_vnd_code").ToString.Trim, "")) & "', " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rej_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rej_reason").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_insp_req").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_status").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_on_behalf").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_wh").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_uom2").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_qty2").ToString.Trim, "0")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(rows.Item("GRD_REJ_RMKS").ToString.Trim)) & ", " &
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _
                                    '"N'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_wh").ToString.Trim, "")) & "', " & _
                                    '"N'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_loc").ToString.Trim, "")) & "', " & _
                            End Select
                            REM **********************

                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                            'If itemSQL <> "" Then
                            '    insertSql = "insert into WMS_GOODSRCV_INSP (" & _
                            '                    "IMP_CODE, STORER_CODE, GR_CODE, GRD_SEQ, GRI_SEQ, " & _
                            '                    "GRI_ITM_CODE, GRI_PACK_KEY, GRI_PALLET_NO, GRI_BATCH_NO, " & _
                            '                    "GRI_SERIAL_NO_LIST, GRI_INSP_QTY, GRI_PASS_QTY, GRI_REJ_QTY, GRI_REJ_REASON, GRI_PHOTO, GRI_STATUS, " & _
                            '                    "SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                            '                "select d.IMP_CODE, d.STORER_CODE, d.GR_CODE, d.GRD_SEQ, row_number() over(partition by D.IMP_CODE, D.STORER_CODE, D.GR_CODE order by d.GRD_SEQ), " & _
                            '                    "d.GRD_ITM_CODE, d.GRD_PACK_KEY, d.GRD_PALLET_NO, d.GRD_BATCH_NO, " & _
                            '                    "null, d.GRD_RCV_QTY, 0, 0, null, null, 'PENDING', " & _
                            '                    "'" & Session("usr_id") & "', Getdate(), '" & Session("usr_id") & "', Getdate() " & _
                            '                "from WMS_GOODSRCV_D d " & _
                            '                "where d.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            '                "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            '                "and d.GR_CODE = '" & gU.dbEncode(nextNo) & "' " & _
                            '                "and d.GRD_ON_BEHALF = 'Y' " & _
                            '                "and d.GRD_INSP_REQ = 'Y' " & _
                            '                "and not exists (" & _
                            '                    "select 1 from WMS_GOODSRCV_INSP i " & _
                            '                    "where d.imp_code = i.imp_code " & _
                            '                    "and d.storer_code = i.storer_code " & _
                            '                    "and d.gr_code = i.gr_code " & _
                            '                    "and d.grd_seq = i.grd_seq) "

                            '    gDB.amendData(insertSql, gConn, transaction)
                            'End If

                        Next
                    End If

                    For Each rows As DataRow In pa_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"

                                itemSQL = "insert into WMS_GOODSRCV_PA " &
                                        "(IMP_CODE, STORER_CODE, GR_CODE, GRA_SEQ, GRA_DISP_SEQ, GRA_ITM_CODE, GRA_PACK_KEY, " &
                                        "GRA_WH, GRA_LOC, GRA_PA_QTY, GRA_SUG_QTY, GRA_SPLIT_FR, GRA_PALLET_NO, GRA_REF_NO, GRA_BATCH_NO, " &
                                        "GRA_EXPIRY_DATE, GRA_MANU_DATE, GRA_PA_LIST_NO, " &
                                        "GRA_REJ_QTY, GRA_REJ_REASON, GRA_REMARK, " &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_disp_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_itm_code").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pack_key").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_wh").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_loc").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_pa_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_sug_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_split_fr").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pallet_no").ToString.Trim, "000")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_ref_no").ToString.Trim, "")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_MANU_DATE").ToString.Trim, ""))) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRA_PA_LIST_NO").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_rej_qty").ToString.Trim, "0")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_rej_reason").ToString.Trim, "")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_remark").ToString.Trim, "")) & "', " &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _
                        End Select
                        REM **********************

                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    serial_seq = 0

                    For Each rows As DataRow In sn_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"

                                serial_seq += 1

                                itemSQL = "insert into WMS_GOODSRCV_D_S (" & _
                                            "IMP_CODE, STORER_CODE, GR_CODE, " & _
                                            "GRS_ITM_CODE, GRS_PACK_KEY, GRS_BATCH_NO, GRS_PALLET_NO, GRS_SEQ, GRS_SERIAL_NO, GRS_LOC, " & _
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values (" & _
                                            "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_ITM_CODE").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PACK_KEY").ToString.Trim, "")) & "', " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("GRS_BATCH_NO").ToString.Trim, ""))) & ", " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PALLET_NO").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(CStr(serial_seq)) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_SERIAL_NO").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_LOC").ToString.Trim, "")) & "', " & _
                                            "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                gDB.amendData(itemSQL, gConn, transaction)
                        End Select
                        REM **********************
                    Next

                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update WMS_GOODSRCV set " &
                                    "GR_STATUS = '" & gU.dbEncode(GR_STATUS.Text.Trim) & "', " &
                                    "GR_DATE = " & gU.convdbDate(gU.dbEncode(GR_DATE.Text.Trim)) & ", " &
                                    "GR_RCV_BY = '" & gU.dbEncode(GR_RCV_BY.Text.Trim) & "', " &
                                    "GR_DOC_NO = '" & gU.dbEncode(GR_DOC_NO.Value.Trim) & "', " &
                                    "GR_DOC_TYPE = '" & gU.dbEncode(GR_DOC_TYPE.Value.Trim) & "', " &
                                    "PO_DATE = " & gU.convdbDate(gU.dbEncode(PO_DATE.Text.Trim)) & ", " &
                                    "VND_CODE = '" & gU.dbEncode(dsp_VND_CODE.Text.Trim) & "', " &
                                    "VND_NAME = '" & gU.dbEncode(VND_NAME.Value) & "', " &
                                    "GR_VND_DNREF = '" & gU.dbEncode(GR_VND_DNREF.Text.Trim) & "', " &
                                    "GR_REM = '" & gU.dbEncode(GR_REM.Text.Trim) & "', " &
                                    "GR_DESTINATION = '" & gU.dbEncode(GR_DESTINATION.Text.Trim) & "', " &
                                    "GR_TOT_PALLET = '" & gU.dbEncode(gU.decodeNullOrEmpty(GR_TOT_PALLET.Text.Trim, "0")) & "', " &
                                    "GR_TRACK_NO = '" & gU.dbEncode(txt_TRK_NO.Text.Trim) & "', " &
                                    "GR_REF_NO = '" & gU.dbEncode(GR_REF_NO.Value) & "', " &
                                    "GR_HAWB = '" & gU.dbEncode(GR_HAWB.Text.Trim) & "', " &
                                    "GR_INV_NO = '" & gU.dbEncode(GR_INV_NO.Text.Trim) & "', " &
                                    "PRJ_CODE = '" & gU.dbEncode(PRJ_CODE.SelectedValue) & "', " &
                                    "GR_TRANS_TYPE = '" & gU.dbEncode(GR_TRANS_TYPE.SelectedValue) & "'," &
                                     "GR_CAT = '" & gU.dbEncode(gr_cat.SelectedValue) & "'," &
                                    "GR_EDI_PO_NO = '" & gU.dbEncode(GR_EDI_PO_NO.Text.Trim) & "', " &
                                    "GR_WH_CODE = '" & gU.dbEncode(GR_WH_CODE.Text.Trim) & "', " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        reOrderGRD(dt)
                        uiFun.reOrderDetails(dt, "grd_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""

                            grd_batch_no = ""
                            grd_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("grd_batch_no").ToString.Trim, gU.decodeNull(rows.Item("GRD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("GRD_EXPIRY_DATE").ToString.Trim, ""))

                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    If grd_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(grd_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSqL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                                        "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                        "'" & gU.dbEncode(grd_batch_no) & "'," & _
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSqL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into WMS_GOODSRCV_D " &
                                            "(imp_code, storer_code, gr_code, grd_seq, grd_disp_seq, grd_pack, grd_pack_no, grd_pallet_no, " &
                                             "grd_carton_no, grd_batch_no, grd_ref_no, grd_itm_code, grd_brand, grd_series, " &
                                             "grd_model, grd_itm_name, grd_po_qty, grd_rcv_qty, grd_os_qty, " &
                                             "grd_uom, grd_width, grd_length, grd_height, grd_kg, grd_cbm, " &
                                             "grd_destination, grd_stocktype, grd_pack_key, grd_pcs_per_uom, grd_tot_pcs, grd_ref_seq, " &
                                             "grd_no_of_carton, grd_pcs_per_carton, grd_vnd_code, grd_expiry_date, grd_manu_date, " &
                                             "grd_rej_qty, grd_rej_reason, grd_insp_req, " &
                                             "grd_status, grd_on_behalf, grd_wh, grd_uom2, grd_qty2,GRD_REJ_RMKS, " &
                                             "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                            "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(GR_CODE.Text) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_disp_seq").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pallet_no").ToString.Trim, "000")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_carton_no").ToString.Trim, "")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(grd_batch_no)) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_ref_no").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_brand").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_series").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_model").ToString.Trim, "")) & "', " &
                                             "N'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_name").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_po_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rcv_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_os_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_uom").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_width").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_length").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_height").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_kg").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_cbm").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_destination").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_stocktype").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_uom").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_tot_pcs").ToString.Trim, "0")) & "', " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_ref_seq").ToString.Trim, "null")) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_no_of_carton").ToString.Trim, "0")) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_carton").ToString.Trim, "0")) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_vnd_code").ToString.Trim, "")) & "', " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rej_qty").ToString.Trim, "0")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rej_reason").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_insp_req").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_status").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_on_behalf").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_wh").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_uom2").ToString.Trim, "")) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_qty2").ToString.Trim, "0")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("GRD_REJ_RMKS").ToString.Trim, ""))) & ", " &
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from WMS_GOODSRCV_D " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                            "and GRD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    If grd_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(grd_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSqL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                                        "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                        "'" & gU.dbEncode(grd_batch_no) & "'," & _
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSqL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "update WMS_GOODSRCV_D set " &
                                                "GRD_DISP_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_disp_seq").ToString.Trim, "")) & "', " &
                                                "GRD_PACK = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack").ToString.Trim, "")) & "', " &
                                                "GRD_PACK_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_no").ToString.Trim, "")) & "', " &
                                                "GRD_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pallet_no").ToString.Trim, "000")) & "', " &
                                                "GRD_CARTON_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_carton_no").ToString.Trim, "")) & "', " &
                                                "GRD_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(grd_batch_no)) & ", " &
                                                "GRD_REF_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_ref_no").ToString.Trim, "")) & "', " &
                                                "GRD_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, "")) & "', " &
                                                "GRD_BRAND = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_brand").ToString.Trim, "")) & "', " &
                                                "GRD_SERIES = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_series").ToString.Trim, "")) & "', " &
                                                "GRD_MODEL = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_model").ToString.Trim, "")) & "', " &
                                                "GRD_ITM_NAME = N'" & gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_name").ToString.Trim, "")) & "', " &
                                                "GRD_PO_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_po_qty").ToString.Trim, "0")) & "', " &
                                                "GRD_RCV_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rcv_qty").ToString.Trim, "0")) & "', " &
                                                "GRD_OS_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_os_qty").ToString.Trim, "0")) & "', " &
                                                "GRD_UOM = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_uom").ToString.Trim, "")) & "', " &
                                                "GRD_WIDTH = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_width").ToString.Trim, "0")) & "', " &
                                                "GRD_LENGTH = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_length").ToString.Trim, "0")) & "', " &
                                                "GRD_HEIGHT = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_height").ToString.Trim, "0")) & "', " &
                                                "GRD_KG = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_kg").ToString.Trim, "0")) & "', " &
                                                "GRD_CBM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_cbm").ToString.Trim, "0")) & "', " &
                                                "GRD_DESTINATION = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_destination").ToString.Trim, "")) & "', " &
                                                "GRD_STOCKTYPE = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_stocktype").ToString.Trim, "")) & "', " &
                                                "GRD_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, "")) & "', " &
                                                "GRD_PCS_PER_UOM = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_uom").ToString.Trim, "0")) & "', " &
                                                "GRD_TOT_PCS = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_tot_pcs").ToString.Trim, "0")) & "', " &
                                                "GRD_REF_SEQ = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_ref_seq").ToString.Trim, "null")) & ", " &
                                                "GRD_NO_OF_CARTON = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_no_of_carton").ToString.Trim, "0")) & ", " &
                                                "GRD_PCS_PER_CARTON = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_carton").ToString.Trim, "0")) & ", " &
                                                "grd_vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_vnd_code").ToString.Trim, "")) & "', " &
                                                "GRD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                "GRD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                "GRD_REJ_QTY = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("GRD_REJ_QTY").ToString.Trim, "0")) & ", " &
                                                "GRD_REJ_REASON = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_REJ_REASON").ToString.Trim, "")) & "', " &
                                                "GRD_INSP_REQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_INSP_REQ").ToString.Trim, "")) & "', " &
                                                "GRD_STATUS = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_STATUS").ToString.Trim, "")) & "', " &
                                                "GRD_ON_BEHALF = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_ON_BEHALF").ToString.Trim, "")) & "', " &
                                                "GRD_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_WH").ToString.Trim, "")) & "', " &
                                                "GRD_UOM2 = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRD_UOM2").ToString.Trim, "")) & "', " &
                                                "GRD_QTY2 = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("GRD_QTY2").ToString.Trim, "0")) & ", " &
                                                "GRD_REJ_RMKS=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("GRD_REJ_RMKS").ToString.Trim, ""))) & ", " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " &
                                            "and GRD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("grd_seq").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If

                    If Session("_M_IB_GR_TMP_pa_del_list") <> "" Then
                        Dim delSeqArray As String()

                        delSeqArray = Split(Session("_M_IB_GR_TMP_pa_del_list"), ", ")

                        For j = 0 To UBound(delSeqArray)
                            itemSQL = "delete from WMS_GOODSRCV_PA " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                        "and GRA_SEQ = '" & gU.dbEncode(delSeqArray(j).Trim) & "' "

                            gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If

                    For Each rows As DataRow In pa_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"
                                itemSQL = "insert into WMS_GOODSRCV_PA " &
                                        "(IMP_CODE, STORER_CODE, GR_CODE, GRA_SEQ, GRA_DISP_SEQ, GRA_ITM_CODE, GRA_PACK_KEY, " &
                                        "GRA_WH, GRA_LOC, GRA_PA_QTY, GRA_SUG_QTY, GRA_SPLIT_FR, GRA_PALLET_NO, GRA_REF_NO, GRA_BATCH_NO, " &
                                        "GRA_EXPIRY_DATE, GRA_MANU_DATE, GRA_PA_LIST_NO," &
                                        "GRA_REJ_QTY, GRA_REJ_REASON, GRA_REMARK, " &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(GR_CODE.Text) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_disp_seq").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_itm_code").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pack_key").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_wh").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_loc").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_pa_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_sug_qty").ToString.Trim, "0")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_split_fr").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pallet_no").ToString.Trim, "000")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_ref_no").ToString.Trim, "")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_MANU_DATE").ToString.Trim, ""))) & ", " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRA_PA_LIST_NO").ToString.Trim, "")) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_rej_qty").ToString.Trim, "0")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_rej_reason").ToString.Trim, "")) & "', " &
                                         "N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_remark").ToString.Trim, "")) & "', " &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                            Case "D"
                                itemSQL = "delete from WMS_GOODSRCV_PA " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                        "and GRA_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_seq").ToString.Trim, "")) & "' "
                            Case Else
                                itemSQL = "update WMS_GOODSRCV_PA set " &
                                            "GRA_DISP_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_disp_seq").ToString.Trim, "")) & "', " &
                                            "GRA_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_itm_code").ToString.Trim, "")) & "', " &
                                            "GRA_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pack_key").ToString.Trim, "")) & "', " &
                                            "GRA_PA_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_pa_qty").ToString.Trim, "0")) & "', " &
                                            "GRA_SUG_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_sug_qty").ToString.Trim, "0")) & "', " &
                                            "GRA_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_wh").ToString.Trim, "")) & "', " &
                                            "GRA_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_loc").ToString.Trim, "")) & "', " &
                                            "GRA_SPLIT_FR = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_split_fr").ToString.Trim, "")) & "', " &
                                            "GRA_PALLET_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_pallet_no").ToString.Trim, "000")) & "', " &
                                            "GRA_REF_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_ref_no").ToString.Trim, "")) & "', " &
                                            "GRA_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, ""))) & ", " &
                                            "GRA_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                            "GRA_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("GRA_MANU_DATE").ToString.Trim, ""))) & ", " &
                                            "GRA_PA_LIST_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRA_PA_LIST_NO").ToString.Trim, "")) & "', " &
                                            "GRA_REJ_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gra_rej_qty").ToString.Trim, "0")) & "', " &
                                            "GRA_REJ_REASON = N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_rej_reason").ToString.Trim, "")) & "', " &
                                            "GRA_REMARK = N'" & gU.dbEncode(gU.decodeNull(rows.Item("gra_remark").ToString.Trim, "")) & "', " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " &
                                        "and GRA_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("gra_seq").ToString.Trim, "")) & "'"
                        End Select
                        REM **********************

                        'Response.Write(itemSQL)
                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                    itemSQL = "select isnull(max(convert(int, grs_seq)), 0) as value " & _
                              "from WMS_GOODSRCV_D_S " & _
                              "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                              "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                              "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

                    serial_seq = CInt(DB.getValueFromSQL(itemSQL, gConn, transaction))

                    For Each rows As DataRow In sn_dt.Rows
                        itemSQL = ""
                        REM **********************
                        REM Modify Here
                        Select Case rows.Item("mFlag")
                            Case "N"

                                serial_seq += 1

                                itemSQL = "insert into WMS_GOODSRCV_D_S (" & _
                                            "IMP_CODE, STORER_CODE, GR_CODE, " & _
                                            "GRS_ITM_CODE, GRS_PACK_KEY, GRS_BATCH_NO, GRS_PALLET_NO, GRS_SEQ, GRS_SERIAL_NO, GRS_LOC, " & _
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values (" & _
                                            "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(GR_CODE.Text.Trim) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_ITM_CODE").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PACK_KEY").ToString.Trim, "")) & "', " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("GRS_BATCH_NO").ToString.Trim, ""))) & ", " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PALLET_NO").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(CStr(serial_seq)) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_SERIAL_NO").ToString.Trim, "")) & "', " & _
                                            "'" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_LOC").ToString.Trim, "")) & "', " & _
                                            "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                            Case "U"

                                itemSQL = "update WMS_GOODSRCV_D_S " & _
                                          "set GRS_SERIAL_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_SERIAL_NO").ToString.Trim, "")) & "', " & _
                                              "GRS_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_LOC").ToString.Trim, "")) & "' " & _
                                          "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                          "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                          "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                          "and GRS_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_ITM_CODE").ToString.Trim, "")) & "' " & _
                                          "and GRS_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PACK_KEY").ToString.Trim, "")) & "' " & _
                                          "and GRS_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_SEQ").ToString.Trim, "")) & "' "

                            Case "D"

                                itemSQL = "delete from WMS_GOODSRCV_D_S " & _
                                          "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                          "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                          "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                          "and GRS_ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_ITM_CODE").ToString.Trim, "")) & "' " & _
                                          "and GRS_PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_PACK_KEY").ToString.Trim, "")) & "' " & _
                                          "and GRS_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("GRS_SEQ").ToString.Trim, "")) & "' "

                        End Select
                        REM **********************

                        If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                    Next

                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)


                'Create new RO for onbehalf item
                If sFlag = "INTRANS" Then

                    selectSql = "select distinct grd_wh " & _
                                "from wms_goodsrcv_d " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                "and GRD_ON_BEHALF = 'Y' " & _
                                "and GRD_STATUS <> 'INTRANS' "

                    tmpDt = gDB.getDataTable(selectSql, gConn, transaction)


                    For i = 0 To tmpDt.Rows.Count - 1
                        newDocNo = DB.getDocNo("RO", gConn, transaction)

                        newTrackNo = DB.getDocNo("TRACKNO", gConn, transaction)

                        'Insert header row
                        insertSql = "insert into wms_replenish (" & _
                                        "ro_code, ro_track_no, " & _
                                        "imp_code, storer_code, ro_status,  " & _
                                        "ro_edi_po_no, ro_rev_no, ro_vnd_code, ro_wh_code, " & _
                                        "ro_date, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                    "select '" & gU.dbEncode(newDocNo) & "', '" & gU.dbEncode(newTrackNo) & "', " & _
                                        "g.IMP_CODE, g.STORER_CODE, 'NEW', " & _
                                        "g.GR_EDI_PO_NO + '-IT-' + '" & gU.dbEncode(tmpDt.Rows(i).Item("grd_wh").ToString.Trim) & "' , null, g.VND_CODE, '" & tmpDt.Rows(i).Item("grd_wh").ToString.Trim & "', " & _
                                        "g.PO_DATE, " & _
                                        "'" & gU.dbEncode(Session("usr_id")) & "', Getdate(),'" & gU.dbEncode(Session("usr_id")) & "', Getdate() " & _
                                    "from wms_goodsrcv g " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

                        updtRow = gDB.amendData(insertSql, gConn, transaction)

                        'Insert detail row
                        insertSql = "insert into wms_replenish_d ( " & _
                                        "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " & _
                                        "rod_pallet_no, rod_batch_no, rod_itm_code, rod_pack_key, rod_itm_name, " & _
                                        "rod_qty, rod_uom, rod_pcs_per_uom, " & _
                                        "rod_tot_pcs, " & _
                                        "rod_date_req, rod_sch_deli, " & _
                                        "rod_insp_req, rod_curr_rate, " & _
                                        "rod_length, rod_width, rod_height, rod_kg, rod_cbm, " & _
                                        "rod_wh_code, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                    "select '" & gU.dbEncode(newDocNo) & "', d.IMP_CODE, d.STORER_CODE, row_number() over (order by d.GRD_SEQ), row_number() over (order by d.GRD_SEQ), " & _
                                        "d.GRD_PALLET_NO, d.GRD_BATCH_NO, d.GRD_ITM_CODE, d.GRD_PACK_KEY, d.GRD_ITM_NAME, " & _
                                        "d.GRD_RCV_QTY, d.GRD_UOM, d.GRD_PCS_PER_UOM, " & _
                                        "d.GRD_TOT_PCS, " & _
                                        "null, null, " & _
                                        "d.GRD_INSP_REQ, null, " & _
                                        "d.GRD_LENGTH, d.GRD_WIDTH, d.GRD_HEIGHT, d.GRD_KG, d.GRD_CBM, " & _
                                        "d.GRD_WH, " & _
                                        "'" & gU.dbEncode(Session("usr_id")) & "', Getdate(),'" & gU.dbEncode(Session("usr_id")) & "', Getdate() " & _
                                    "from wms_goodsrcv_d d " & _
                                    "where d.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and d.GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                    "and d.GRD_ON_BEHALF = 'Y' " & _
                                    "and d.GRD_STATUS <> 'INTRANS' " & _
                                    "and d.GRD_WH = '" & gU.dbEncode(tmpDt.Rows(i).Item("grd_wh").ToString.Trim) & "' "

                        updtRow = gDB.amendData(insertSql, gConn, transaction)
                    Next

                    updateSqL = "update WMS_GOODSRCV_D " & _
                                "set GRD_STATUS = 'INTRANS' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                "and GRD_ON_BEHALF = 'Y' " & _
                                "and GRD_STATUS <> 'INTRANS' "

                    updtRow = gDB.amendData(updateSqL, gConn, transaction)

                    updateSqL = "update WMS_GOODSRCV " & _
                                "set GR_STATUS = 'INTRANS' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                "and not exists (" & _
                                    "select 1 from WMS_GOODSRCV_D d " & _
                                    "where d.IMP_CODE = WMS_GOODSRCV.IMP_CODE " & _
                                    "and d.STORER_CODE = WMS_GOODSRCV.STORER_CODE " & _
                                    "and d.GR_CODE = WMS_GOODSRCV.GR_CODE " & _
                                    "and isnull(GRD_ON_BEHALF, 'N') <> 'Y') "

                    updtRow = gDB.amendData(updateSqL, gConn, transaction)

                ElseIf sFlag = "RELEASE" Then

                    updateSqL = "update WMS_GOODSRCV " & _
                                "set GR_STATUS = 'WPA' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSqL, gConn, transaction)

                    'updateSqL = "update WMS_GOODSRCV_D " & _
                    '            "set GRD_STATUS = 'WPA' " & _
                    '            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    '            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    '            "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

                    'gDB.amendData(updateSqL, gConn, transaction)

                    updateSqL = "update WMS_GOODSRCV_PA " & _
                                "set GRA_STATUS = 'WPA' " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

                    gDB.amendData(updateSqL, gConn, transaction)

                Else

                    If GR_STATUS.Text = "WPA" Then
                        updateSqL = "update WMS_GOODSRCV_PA " & _
                                    "set GRA_STATUS = 'WPA' " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                                    "and isnull(GRA_STATUS, 'NEW') = 'NEW' " & _
                                    "and exists (" & _
                                        "select 1 from WMS_GOODSRCV g " & _
                                        "where g.IMP_CODE = WMS_GOODSRCV_PA.IMP_CODE " & _
                                        "and g.STORER_CODE = WMS_GOODSRCV_PA.STORER_CODE " & _
                                        "and g.GR_CODE = WMS_GOODSRCV_PA.GR_CODE " & _
                                        "and g.GR_STATUS = 'WPA') "

                        gDB.amendData(updateSqL, gConn, transaction)
                    End If

                End If

                updatePAParentChild(gConn, transaction)

                For Each rows As DataRow In dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                dt.AcceptChanges()

                For Each rows As DataRow In pa_dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                pa_dt.AcceptChanges()

                For Each rows As DataRow In sn_dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                sn_dt.AcceptChanges()

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    GR_CODE.Text = nextNo
                    GR_CODE_HF.Value = nextNo
                    ViewState("GR_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    GR_CODE.ForeColor = Drawing.Color.Black
                    GR_CODE.Font.Size = 10
                    REM **********************
                End If

                If sFlag = "" Then
                    reloadPage("Record has been saved successfully!")
                End If

                'If sFlag = "" Then uiFun.displayMsgNew(updtPnlAlert, "1007", "", Session("gLang"))
                'ViewState("n_cur_seq") = ""
                'Call BindGV()

                'Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
                'updtPnlLinkBar.Update()

                save = True

            Catch ex As Exception
                'transaction.Rollback()
                'Response.Write(ex.Message)
                'uiFun.displayMsgNew(updtPnlAlert, "1008", "", Session("gLang"))

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

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim grCode, storerCode As String
        Dim insp_dt As DataTable

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("GR_CODE") <> "" Then
            'pk_code = Session("GR_CODE")
            grCode = ViewState("GR_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            'pk_code = Request("GR_CODE")
            grCode = Server.UrlDecode(Request("GR_CODE"))
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
            GR_CODE.ForeColor = Drawing.Color.Red
            GR_STATUS.Text = "NEW"
            IMP_CODE.Value = Session("IMP_CODE")
            ViewState("STORER_CODE") = STORER_CODE.SelectedValue
            btnAttach.Visible = False
            btnNOTE.visible = False
            REM **********************
        Else
            btnAttach.Visible = True
            btnNOTE.visible = True
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "SELECT " &
                            "IMP_CODE, " &
                            "STORER_CODE, " &
                            "GR_CODE, " &
                            "PRJ_CODE, " &
                            "GR_STATUS, " &
                            "GR_TYPE, " &
                            "CONVERT(varchar,GR_DATE," & DDFORMAT & ") as GR_DATE, " &
                            "GR_RCV_BY, " &
                            "PO_CODE, " &
                            "GR_DOC_TYPE, " &
                            "GR_DOC_NO, " &
                            "convert(varchar, PO_DATE," & DDFORMAT & ") as PO_DATE, " &
                            "GR_DESTINATION, " &
                            "VND_CODE, " &
                            "VND_NAME, " &
                            "GR_VND_DNREF, " &
                            "GR_TOT_PALLET, " &
                            "GR_REM, " &
                            "GR_TRACK_NO, " &
                            "GR_REF_NO, " &
                            "convert(varchar, GR_POSTED_DATE," & DDFORMAT & ") as GR_POSTED_DATE, " &
                            "GR_POSTED_BY, " &
                            "SYS_LUB, " &
                            "SYS_LUD, " &
                            "SYS_CD, " &
                            "SYS_CB, " &
                            "GR_HAWB, " &
                            "GR_INV_NO, " &
                            "GR_TRANS_TYPE, " &
                            "GR_CAT, " &
                            "GR_EDI_PO_NO, " &
                            "GR_WH_CODE " &
                        "from WMS_GOODSRCV " &
                        "WHERE WMS_GOODSRCV.GR_CODE = '" & gU.dbEncode(grCode) & "' " &
                        "AND WMS_GOODSRCV.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                        "AND WMS_GOODSRCV.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                GR_CODE.Text = dt.Rows(0).Item("GR_CODE").ToString
                GR_CODE_HF.Value = dt.Rows(0).Item("GR_CODE").ToString
                GR_STATUS.Text = dt.Rows(0).Item("GR_STATUS").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString

                ViewState("STORER_CODE") = dt.Rows(0).Item("STORER_CODE").ToString

                GR_DATE.Text = dt.Rows(0).Item("GR_DATE").ToString
                GR_RCV_BY.Text = dt.Rows(0).Item("GR_RCV_BY").ToString
                PRJ_CODE.SelectedValue = dt.Rows(0).Item("PRJ_CODE").ToString
                'lbl_PRJ_NAME.Text = dt.Rows(0).Item("PRJ_NAME").ToString
                'New Field
                GR_HAWB.Text = dt.Rows(0).Item("GR_HAWB").ToString
                GR_INV_NO.Text = dt.Rows(0).Item("GR_INV_NO").ToString
                txt_TRK_NO.Text = dt.Rows(0).Item("GR_TRACK_NO").ToString
                'XXXXXXXXXX
                GR_DOC_TYPE.Value = dt.Rows(0).Item("GR_DOC_TYPE").ToString
                txt_GR_DOC_TYPE.Text = dt.Rows(0).Item("GR_DOC_TYPE").ToString
                GR_DOC_NO.Value = dt.Rows(0).Item("GR_DOC_NO").ToString
                txt_GR_DOC_NO.Text = dt.Rows(0).Item("GR_DOC_NO").ToString
                PO_DATE.Text = dt.Rows(0).Item("PO_DATE").ToString
                VND_CODE.Value = dt.Rows(0).Item("VND_CODE").ToString
                dsp_VND_CODE.Text = dt.Rows(0).Item("VND_CODE").ToString
                VND_NAME.Value = dt.Rows(0).Item("VND_NAME").ToString
                dsp_VND_NAME.Text = dt.Rows(0).Item("VND_NAME").ToString
                GR_VND_DNREF.Text = dt.Rows(0).Item("GR_VND_DNREF").ToString
                GR_DESTINATION.Text = dt.Rows(0).Item("GR_DESTINATION").ToString
                GR_TOT_PALLET.Text = dt.Rows(0).Item("GR_TOT_PALLET").ToString
                GR_REM.Text = dt.Rows(0).Item("GR_REM").ToString
                GR_TRACK_NO.Value = dt.Rows(0).Item("GR_TRACK_NO").ToString
                GR_REF_NO.Value = dt.Rows(0).Item("GR_REF_NO").ToString

                GR_EDI_PO_NO.Text = dt.Rows(0).Item("GR_EDI_PO_NO").ToString

                GR_WH_CODE.Text = dt.Rows(0).Item("GR_WH_CODE").ToString
                GR_WH_CODE.Attributes.Add("readonly", "")

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                GR_TRANS_TYPE.SelectedValue = dt.Rows(0).Item("GR_TRANS_TYPE").ToString.Trim
                gr_cat.SelectedValue = dt.Rows(0).Item("GR_CAT").ToString.Trim
                'lbl_PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("PRJ_CODE").ToString) & "' ")
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here

        REM Generate Data Table from Detail 2

        SQLString = "SELECT d.IMP_CODE, d.STORER_CODE, d.GR_CODE, d.GRA_SEQ, d.GRA_DISP_SEQ, d.GRA_ITM_CODE, i.ITM_NAME as DSP_ITM_NAME, d.GRA_PACK_KEY, " & _
                        "d.GRA_WH, d.GRA_LOC, d.GRA_PA_QTY, d.GRA_SUG_QTY, d.GRA_SPLIT_FR, d.GRA_PALLET_NO, d.GRA_REF_NO, d.GRA_BATCH_NO, i.ITM_SKU_NO, " & _
                        "d.GRA_REJ_QTY, d.GRA_REJ_REASON, d.GRA_REMARK, " & _
                        "convert(varchar, d.GRA_EXPIRY_DATE, " & gU.dbEncode(DDFORMAT) & ") as GRA_EXPIRY_DATE, " & _
                        "convert(varchar, d.GRA_MANU_DATE, " & gU.dbEncode(DDFORMAT) & ") as GRA_MANU_DATE,GRA_PA_LIST_NO, " & _
                        "'U' as mFlag, ' ' as chk_loc " & _
                    "from WMS_GOODSRCV_PA d " & _
                    "left outer join WMS_ITEM i " & _
                    "on d.IMP_CODE = i.IMP_CODE " & _
                    "and d.STORER_CODE = i.STORER_CODE " & _
                    "and d.GRA_ITM_CODE = i.ITM_CODE " & _
                    "and d.GRA_PACK_KEY = i.PACK_KEY " & _
                    "where d.GR_CODE = '" & gU.dbEncode(grCode) & "' " & _
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " & _
                    "and d.IMP_CODE = '" & Session("IMP_CODE") & "' "


        SQLString = SQLString & " order by d.GRA_DISP_SEQ"
        REM **********************
        Session("_M_IB_GR_TMP_pa_dt") = gDB.getDataTable(SQLString)
        REM **********************

        SQLString = "select MAX(convert(int,GRA_SEQ)) from WMS_GOODSRCV_PA " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

        Session("_M_IB_GR_TMP_pa_seq") = DB.getValueFromSQL(SQLString)

        If Session("_M_IB_GR_TMP_pa_seq") = "" Then
            Session("_M_IB_GR_TMP_pa_seq") = "0"
        End If

        Session("_M_IB_GR_TMP_pa_del_list") = ""

        REM Generate Data Table from Detail
        'SQLString = "SELECT d.IMP_CODE, d.STORER_CODE, d.GR_CODE, d.GRD_SEQ, d.GRD_PACK, d.GRD_PACK_NO, d.GRD_ITM_CODE, d.GRD_ITM_NAME, " & _
        '                "d.GRD_WH, d.GRD_LOC, d.GRD_PO_QTY, d.GRD_RCV_QTY, d.GRD_OS_QTY, 'U' as mFlag " & _
        '            "from WMS_GOODSRCV_D d " & _
        '            "where d.GR_CODE = '" & gU.dbEncode(grCode) & "' " & _
        '            "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " & _
        '            "and d.IMP_CODE = '" & Session("IMP_CODE") & "'"

        SQLString = "SELECT d.GRD_DISP_SEQ, d.GRD_ITM_CODE, IsNUll(d.GRD_ITM_NAME,I.ITM_NAME)GRD_ITM_NAME, d.GRD_PACK, d.GRD_PACK_NO, " &
                        "d.GRD_PO_QTY, d.GRD_RCV_QTY, d.GRD_OS_QTY, d.GRD_WH, d.GRD_LOC, " &
                        "d.GRD_SEQ, d.GRD_REF_SEQ, d.GRD_PALLET_NO, d.GRD_CARTON_NO,  d.GRD_BATCH_NO, d.GRD_REF_NO, d.GRD_BRAND, d.GRD_SERIES, " &
                        "d.GRD_MODEL, d.GRD_UOM, d.GRD_LENGTH, d.GRD_WIDTH, d.GRD_HEIGHT, d.GRD_KG, d.GRD_CBM, " &
                        "d.GRD_DESTINATION, d.GRD_STOCKTYPE,d.GRD_PACK_KEY, d.GRD_PCS_PER_UOM, d.GRD_TOT_PCS, " &
                        "d.GRD_NO_OF_CARTON, d.GRD_PCS_PER_CARTON, d.grd_vnd_code, " &
                        "Convert(varchar,d.GRD_EXPIRY_DATE," & DDFORMAT & ") as GRD_EXPIRY_DATE, " &
                        "Convert(varchar, d.GRD_MANU_DATE," & DDFORMAT & ") as GRD_MANU_DATE, " &
                        "d.GRD_REJ_QTY, d.GRD_REJ_REASON, d.GRD_INSP_REQ, d.GRD_STATUS, " &
                        "d.GRD_ON_BEHALF, d.grd_uom2, d.grd_qty2,d.GRD_REJ_RMKS, " &
                        "i.ITM_SKU_NO, i.ITM_DESC, i.ITM_SERIAL_NO_YN, i.itm_type, " &
                        "'U' as mFlag, '0' as cb_select " &
                    "from WMS_GOODSRCV_D d left outer join WMS_ITEM i " &
                    "on d.IMP_CODE = i.IMP_CODE " &
                    "and d.STORER_CODE = i.STORER_CODE " &
                    "and d.GRD_ITM_CODE = i.ITM_CODE " &
                    "and d.GRD_PACK_KEY = i.PACK_KEY " &
                    "where d.GR_CODE = '" & gU.dbEncode(grCode) & "' " &
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " &
                    "and d.IMP_CODE = '" & Session("IMP_CODE") & "' "


        SQLString = SQLString & " order by convert(int, d.GRD_SEQ)"
        'SQLString = SQLString & " order by to_number(ISNULL(d.GRD_REF_SEQ, d.GRD_SEQ)), d.GRD_SEQ, d.GRD_DISP_SEQ"
        'SQLString = SQLString & " order by d.GRD_SEQ, to_number(ISNULL(d.GRD_REF_SEQ, d.GRD_SEQ)),  d.GRD_DISP_SEQ"
        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        Session("dt") = dt

        GridView1.DataBind()

        SQLString = "select s.IMP_CODE, s.STORER_CODE, s.GR_CODE, s.GRS_ITM_CODE, s.GRS_PACK_KEY, s.GRS_BATCH_NO, s.GRS_PALLET_NO, s.GRS_SEQ, s.GRS_SERIAL_NO, s.GRS_LOC, " & _
                        "s.GRS_DATE_IN, s.GRS_CO_CODE, s.GRS_DO_CODE, s.GRS_WARR_STDATE, s.GRS_WARR_TYPE, " & _
                        "s.GRS_WARR_PERIOD, s.GRS_WARR_EXPDATE, s.CUST_NAME, s.CONS_NO, s.CONS_NAME, s.CONS_TEL, " & _
                        "s.CONS_ADDR, s.CONS_CTRY, " & _
                        "'U' as mFlag " & _
                    "from wms_goodsrcv_d_s s " & _
                    "where s.gr_code = '" & gU.dbEncode(grCode) & "' " & _
                    "and s.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and s.imp_code = '" & Session("imp_code") & "'"

        '"where d.imp_code = s.imp_code " & _
        '            "and d.storer_code = s.storer_code " & _
        '            "and d.gr_code = s.gr_code " & _
        '            "and d.grd_itm_code = s.grs_itm_code " & _
        '            "and d.grd_pack_key = s.grs_pack_key " & _

        SQLString = SQLString & " order by s.GRS_ITM_CODE, s.GRS_PACK_KEY, s.GRS_BATCH_NO, s.GRS_PALLET_NO, s.GRS_SERIAL_NO, S.GRS_SEQ "

        sn_dt = gDB.getDataTable(SQLString)
        Session("sn_dt") = sn_dt

        REM **********************

        'SQLString = "select i.IMP_CODE, i.STORER_CODE, i.GR_CODE, i.GRD_SEQ, i.GRI_SEQ, i.GRI_ITM_CODE, i.GRI_PACK_KEY, i.GRI_BATCH_NO, i.GRI_SERIAL_NO_LIST, i.GRI_INSP_QTY, " & _
        '                "i.GRI_PASS_QTY, i.GRI_REJ_QTY, i.GRI_REJ_REASON, i.GRI_PHOTO, i.GRI_STATUS, " & _
        '                "t.ITM_NAME, t.ITM_SKU_NO, " & _
        '                "'U' as mFlag " & _
        '            "from WMS_GOODSRCV_INSP i " & _
        '            "left outer join WMS_ITEM t " & _
        '                "on i.IMP_CODE = t.IMP_CODE " & _
        '                "and i.STORER_CODE = t.STORER_CODE " & _
        '                "and i.GRI_ITM_CODE = t.ITM_CODE " & _
        '                "and i.GRI_PACK_KEY = t.PACK_KEY " & _
        '            "where i.gr_code = '" & gU.dbEncode(grCode) & "' " & _
        '            "and i.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
        '            "and i.imp_code = '" & Session("imp_code") & "' " & _
        '            "order by i.GRI_ITM_CODE, i.GRI_PACK_KEY, i.GRI_BATCH_NO, i.GRI_SEQ "

        'insp_dt = gDB.getDataTable(SQLString)
        'Session("gr_insp_dt") = insp_dt

    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_goodsrcv " & _
                     "set gr_status = 'CANCELLED', " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and gr_code = '" & gU.dbEncode(GR_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            'GR_STATUS.Text = "CANCELLED"

            'ar.sec_write = "N"
            'CancelBtn.Visible = False
            'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            'Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
            'updtPnlLinkBar.Update()

            'uiFun.displayMsg(Me, "1011", "", Session("gLang"))

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

    Protected Sub btnInTransit_Click(sender As Object, e As System.EventArgs) Handles btnInTransit.Click
        If save("INTRANS") Then
            reloadPage("On-behalf record has been generated successfully!")
        End If
    End Sub

    Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
        If save("RELEASE") Then
            reloadPage("Put Away has been released to PDA!")
        End If
    End Sub

    'Protected Sub VND_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VND_CODE.SelectedIndexChanged
    '    VND_NAME.Text = DB.getValueFromSQL("select VND_NAME from WMS_VENDOR where VND_CODE = '" & VND_CODE.SelectedValue & "' ")
    'End Sub

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Dim selectSql, updateSql, insertSql As String
        Dim gConn As SqlConnection
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim itmKey As String
        Dim pa_dt, snDt As DataTable
        Dim successFlag As Boolean
        REM Hold Logic
        Dim holddt As DataTable
        Dim coHolddt As DataTable
        Dim checkdt As DataTable
        Dim SQLString As String
        Dim dmgLoc As String

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()

        SQLString = "SELECT RO_STATUS FROM WMS_REPLENISH M " & _
                    "WHERE RO_STATUS = 'CLOSED' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' "
        checkdt = gDB.getDataTable(SQLString)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The PO status is closed, please re-open before post GR!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The PO status is closed, please re-open before post GR!", Session("gLang"))
            End If
            Exit Sub
        End If

        SQLString = "SELECT GR_STATUS FROM wms_goodsrcv M " & _
                    "WHERE GR_STATUS = 'POSTED' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The GR status is closed, please re-open before post GR!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The GR status is closed, please re-open before post GR!", Session("gLang"))
            End If
            Exit Sub
        End If

        If GridView1.Rows.Count > 0 Then
            'For i = 0 To GridView1.Rows.Count - 1
            '    Dim loc As String = CType(GridView1.Rows(i).FindControl("grd_loc"), HiddenField).Value

            '    If loc = "" Then
            '        If Session("gLang") = "E" Then
            '            uiFun.displayMsgNew(updtPnlAlert, "", "Location Cannot Be Empty!", Session("gLang"))
            '        Else
            '            uiFun.displayMsgNew(updtPnlAlert, "", "位置不能空白!", Session("gLang"))
            '        End If

            '        Exit Sub
            '    End If
            'Next

            pa_dt = Session("_M_IB_GR_TMP_pa_dt")

            snDt = Session("sn_dt")

            If pa_dt.Rows.Count > 0 Then
                If validateAll("P") Then
                    If save("P") Then
                        gConn = gDB.getConnection()

                        Dim transaction As SqlTransaction
                        transaction = gConn.BeginTransaction()

                        Try

                            SQLString = "select min(loc) as value " & _
                                        "from v_location " & _
                                        "where ar_damage_yn = 'Y' " & _
                                        "and isnull(ar_insp_area, 'N') <> 'Y' " & _
                                        "and wh_code = '" & gU.dbEncode(GR_WH_CODE.Text.Trim) & "' "

                            dmgLoc = DB.getValueFromSQL(SQLString, gConn, transaction)

                            If dmgLoc = "" Then
                                SQLString = "select min(loc) as value " & _
                                            "from v_location " & _
                                            "where ar_damage_yn = 'Y' " & _
                                            "and isnull(ar_insp_area, 'N') <> 'Y' "

                                dmgLoc = DB.getValueFromSQL(SQLString, gConn, transaction)
                            End If

                            qtyDict = New Dictionary(Of String, Double)
                            cbmDict = New Dictionary(Of String, Double)
                            wgtDict = New Dictionary(Of String, Double)

                            For Each rows As DataRow In dt.Rows
                                Dim cbm As Double = gU.decodeEmptyCdbl(rows.Item("grd_length").ToString, 0) * _
                                                        gU.decodeEmptyCdbl(rows.Item("grd_width").ToString, 0) * _
                                                        gU.decodeEmptyCdbl(rows.Item("grd_height").ToString, 0)

                                If cbm <> 0 Then cbm = cbm / 1000000


                                itmKey = gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, "") & "#_#" & _
                                            gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, "") & "#_#" & _
                                            gU.decodeNull(rows.Item("grd_pallet_no").ToString.Trim, "") & "#_#" & _
                                            gU.decodeNull(rows.Item("grd_batch_no").ToString.Trim, "")

                                If qtyDict.ContainsKey(itmKey) Then
                                    qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("grd_rcv_qty").ToString.Trim, 0)
                                    cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(cbm, 0), 14)
                                    wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("GRD_KG").ToString.Trim, 0) * gU.decodeEmptyCdbl(rows.Item("GRD_NO_OF_CARTON").ToString.Trim, 1), 14)
                                Else
                                    qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("grd_rcv_qty").ToString.Trim, 0))
                                    cbmDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(cbm, 0), 14))
                                    wgtDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(rows.Item("GRD_KG").ToString.Trim, 0) * gU.decodeEmptyCdbl(rows.Item("GRD_NO_OF_CARTON").ToString.Trim, 1), 14))
                                End If
                            Next

                            REM Hold Logic
                            SQLString = "select gh.gr_code, grd_itm_code, grd_batch_no, grd_pack_key, grd_pallet_no, grd_rcv_qty " & _
                                        "from wms_goodsrcv gh, wms_goodsrcv_d gd " & _
                                        "where gh.storer_code = gd.storer_code " & _
                                        "and gh.imp_code = gd.imp_code  " & _
                                        "and gh.gr_code = gd.gr_code " & _
                                        "and gh.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                        "and gh.imp_code = '" & Session("imp_code") & "' " & _
                                        "and gh.gr_code = '" & gU.dbEncode(GR_CODE.Text) & "' "
                            SQLString = SQLString & " order by 1, 2"

                            holddt = gDB.getDataTable(SQLString, gConn, transaction)

                            For i As Integer = 0 To holddt.Rows.Count - 1
                                Dim grd_rcv_qty As Integer = 0
                                grd_rcv_qty = gU.decodeEmptyCInt(holddt.Rows(i).Item("grd_rcv_qty"), 0)
                                SQLString = "select hold.co_code, hold.cod_seq, hold.coh_seq, COH_IN_STOCK_QTY, COH_QTY " & _
                                        "from wms_cust_order_hold hold, wms_cust_order_d cd " & _
                                        "where hold.storer_code = cd.storer_code " & _
                                        "and hold.imp_code = cd.imp_code " & _
                                        "and hold.co_code = cd.co_code " & _
                                        "and hold.cod_seq = cd.cod_seq " & _
                                        "and hold.coh_type = 'RO' " & _
                                        "and hold.coh_status = 'HOLD' " & _
                                        "and hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                        "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                        "and hold.ro_code = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                        "and hold.coh_itm_code = '" & gU.dbEncode(holddt.Rows(i).Item("grd_itm_code").ToString) & "' " & _
                                        "and isnull(hold.coh_batch_no, '') = isnull('" & gU.dbEncode(holddt.Rows(i).Item("grd_batch_no").ToString) & "', '') " & _
                                        "and hold.coh_pack_key = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pack_key").ToString) & "' " & _
                                        "and isnull(hold.coh_pallet_no, '000') = isnull('" & gU.dbEncode(holddt.Rows(i).Item("grd_pallet_no").ToString) & "', '000') "

                                '"and cd.cod_itm_code = '" & gU.dbEncode(holddt.Rows(i).Item("grd_itm_code").ToString) & "' " & _
                                '        "and cd.cod_batch_no = '" & gU.dbEncode(holddt.Rows(i).Item("grd_batch_no").ToString) & "' " & _
                                '        "and cd.cod_pack_key = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pack_key").ToString) & "' " & _
                                '        "and cd.cod_pallet_no = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pallet_no").ToString) & "' "

                                SQLString = SQLString & " order by 1, 2, 3"

                                coHolddt = gDB.getDataTable(SQLString, gConn, transaction)

                                For j As Integer = 0 To coHolddt.Rows.Count - 1
                                    If gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_QTY"), 0) <> gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0) Then
                                        Dim COH_IN_STOCK_QTY As Integer = 0
                                        If gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_QTY"), 0) <= gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0) + grd_rcv_qty Then
                                            COH_IN_STOCK_QTY = gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_QTY"), 0)
                                            grd_rcv_qty = grd_rcv_qty - (gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_QTY"), 0) - gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0))
                                        Else
                                            COH_IN_STOCK_QTY = gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0) + grd_rcv_qty
                                        End If
                                        updateSql = "UPDATE wms_cust_order_hold hold SET COH_IN_STOCK_QTY = " & COH_IN_STOCK_QTY & " " & _
                                                    "WHERE hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                                    "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                                    "and hold.co_code = '" & gU.dbEncode(coHolddt.Rows(j).Item("co_code").ToString) & "' " & _
                                                    "and hold.cod_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("cod_seq").ToString) & "' " & _
                                                    "and hold.coh_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("coh_seq").ToString) & "' " & _
                                                    "and hold.coh_type = 'RO' " & _
                                                    "and hold.coh_status = 'HOLD' "
                                        gDB.amendData(updateSql, gConn, transaction)

                                        REM update to hold Stock
                                        updateSql = "UPDATE wms_cust_order_hold hold SET coh_type = 'STOCK' " & _
                                                    "WHERE hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                                    "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                                    "and hold.co_code = '" & gU.dbEncode(coHolddt.Rows(j).Item("co_code").ToString) & "' " & _
                                                    "and hold.cod_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("cod_seq").ToString) & "' " & _
                                                    "and hold.coh_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("coh_seq").ToString) & "' " & _
                                                    "and hold.COH_IN_STOCK_QTY = hold.COH_QTY " & _
                                                    "and hold.coh_type = 'RO' " & _
                                                    "and hold.coh_status = 'HOLD' "
                                        gDB.amendData(updateSql, gConn, transaction)
                                    End If
                                Next
                            Next

                            Dim tempEDate As String = ""
                            Dim tempMDate As String = ""
                            Dim tempvndCode As String = ""
                            Dim tempUOM2 As String = ""
                            Dim tempQty2 As String = ""

                            For Each rows As DataRow In pa_dt.Rows

                                'check if GR is posted or not everytime, this if condition added on 23.01.25
                                SQLString = "SELECT count(*) FROM wms_goodsrcv M " &
                                "WHERE GR_STATUS = 'POSTED' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "
                                If DB.getValueFromSQL(SQLString, gConn, transaction) <> "0" Then
                                    Exit Sub
                                End If

                                If DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) > 0 OrElse DB.decodeDBNull(rows.Item("GRA_REJ_QTY"), 0) > 0 And rows.Item("mFlag").ToString <> "D" Then
                                    dt = Session("dt")
                                    For Each DROWS As DataRow In dt.Rows
                                        If DROWS.Item("GRD_ITM_CODE").ToString.Trim = rows.Item("gra_itm_code").ToString.Trim AndAlso DROWS.Item("GRD_PACK_KEY").ToString.Trim = rows.Item("GRA_PACK_KEY").ToString.Trim AndAlso
                                           DROWS.Item("GRD_PALLET_NO").ToString.Trim = rows.Item("GRA_PALLET_NO").ToString.Trim AndAlso DROWS.Item("GRD_BATCH_NO").ToString.Trim = rows.Item("GRA_BATCH_NO").ToString.Trim Then

                                            tempEDate = DROWS.Item("GRD_EXPIRY_DATE").ToString.Trim
                                            tempMDate = DROWS.Item("GRD_MANU_DATE").ToString.Trim
                                            tempvndCode = DROWS.Item("GRD_VND_CODE").ToString.Trim

                                            tempUOM2 = DROWS.Item("GRD_UOM2").ToString.Trim
                                            tempQty2 = DROWS.Item("GRD_QTY2").ToString.Trim

                                            Exit For
                                        End If
                                    Next

                                    st.IO_SEQ = ""
                                    st.STORER_CODE = STORER_CODE.SelectedValue
                                    st.ITM_CODE = gU.decodeNull(rows.Item("GRA_ITM_CODE").ToString.Trim, "")
                                    st.PACK_KEY = gU.decodeNull(rows.Item("GRA_PACK_KEY").ToString.Trim, "")
                                    st.IO_CUST_CODE = ""
                                    st.IO_WH = gU.decodeNull(rows.Item("GRA_WH").ToString.Trim, "")
                                    st.IO_AREA = ""
                                    st.IO_LOC = gU.decodeNull(rows.Item("GRA_LOC").ToString.Trim, "")
                                    st.IO_DOC = "GR"
                                    st.IO_DOC_ID = GR_CODE.Text.Trim
                                    st.IO_QTY = gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "")
                                    st.PALLET_NO = gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "000")
                                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, "")
                                    st.IO_MANU_DATE = tempMDate
                                    st.IO_EXPIRY_DATE = tempEDate
                                    st.lO_VND_CODE = tempvndCode

                                    itmKey = gU.decodeNull(rows.Item("GRA_ITM_CODE").ToString.Trim, "") & "#_#" &
                                            gU.decodeNull(rows.Item("GRA_PACK_KEY").ToString.Trim, "") & "#_#" &
                                            gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "") & "#_#" &
                                            gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, "")

                                    If qtyDict.ContainsKey(itmKey) Then
                                        Dim tmpCbm As Decimal = 0
                                        Dim tmpwgt As Decimal = 0
                                        Dim tmpqty As Decimal = 0

                                        tmpCbm = cbmDict.Item(itmKey)
                                        tmpwgt = wgtDict.Item(itmKey)
                                        tmpqty = qtyDict.Item(itmKey)

                                        If tmpqty = 0 Then tmpqty = 1

                                        st.IO_CBM = Math.Round(tmpCbm / tmpqty * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                        st.IO_KG = Math.Round(tmpwgt / tmpqty * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                        'st.IO_CBM = Math.Round(tmpCbm * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                        'st.IO_KG = Math.Round(tmpwgt * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                    Else
                                        st.IO_CBM = 0
                                        st.IO_KG = 0
                                    End If

                                    st.IOS_DRUM_LEVEL = "1"
                                    st.IOS_UOM2 = tempUOM2
                                    st.IOS_ORG_QTY2 = tempQty2
                                    st.IOS_QTY2 = tempQty2
                                    st.IOS_SL = "N"
                                    st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                    'Loop to update serial
                                    If snDt IsNot Nothing AndAlso snDt.Rows.Count > 0 Then

                                        For k = 0 To snDt.Rows.Count - 1
                                            If snDt.Rows(k).Item("GRS_ITM_CODE").ToString.Trim = st.ITM_CODE AndAlso
                                                snDt.Rows(k).Item("GRS_PACK_KEY").ToString.Trim = st.PACK_KEY AndAlso
                                                gU.decodeNull(snDt.Rows(k).Item("GRS_PALLET_NO").ToString.Trim, "000") = st.PALLET_NO AndAlso
                                                snDt.Rows(k).Item("GRS_BATCH_NO").ToString.Trim = st.lO_BATCH_NO AndAlso
                                                snDt.Rows(k).Item("GRS_LOC").ToString.Trim = st.IO_LOC AndAlso
                                                snDt.Rows(k).Item("mFlag").ToString.Trim <> "D" Then

                                                st.IOS_DRUM_ID = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                                st.IOS_ORG_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                                st.IOS_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                                st.IO_QTY = 1

                                                st.UpdateStockTrans("IN", gConn, transaction)

                                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                                st.UpdateStockSerialTrans("IN", gConn, transaction)

                                                st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                            End If
                                        Next
                                    Else
                                        st.UpdateStockTrans("IN", gConn, transaction)

                                        st.UpdateStockBalTrans("IN", gConn, transaction)
                                    End If

                                    'Post for rejected item to damage zone
                                    If CDbl(gU.decodeNullOrEmpty(rows.Item("GRA_REJ_QTY").ToString.Trim, "0")) > 0 Then
                                        If dmgLoc = "" Then
                                            Throw New Exception("No damage location!")
                                        Else
                                            st.IO_LOC = dmgLoc

                                            st.IO_QTY = gU.decodeNull(rows.Item("GRA_REJ_QTY").ToString.Trim, "")

                                            st.UpdateStockTrans("IN", gConn, transaction)

                                            st.UpdateStockBalTrans("IN", gConn, transaction)
                                        End If

                                        'Loop to update serial
                                        If snDt IsNot Nothing AndAlso snDt.Rows.Count > 0 Then

                                            For k = 0 To snDt.Rows.Count - 1
                                                If snDt.Rows(k).Item("GRS_ITM_CODE").ToString.Trim = st.ITM_CODE AndAlso
                                                    snDt.Rows(k).Item("GRS_PACK_KEY").ToString.Trim = st.PACK_KEY AndAlso
                                                    gU.decodeNull(snDt.Rows(k).Item("GRS_PALLET_NO").ToString.Trim, "000") = st.PALLET_NO AndAlso
                                                    snDt.Rows(k).Item("GRS_BATCH_NO").ToString.Trim = st.lO_BATCH_NO AndAlso
                                                    snDt.Rows(k).Item("GRS_LOC").ToString.Trim = dmgLocValue AndAlso
                                                    snDt.Rows(k).Item("mFlag").ToString.Trim <> "D" Then

                                                    st.IOS_DRUM_ID = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                                    st.IOS_ORG_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                                    st.IOS_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim

                                                    st.UpdateStockSerialTrans("IN", gConn, transaction)

                                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                                End If
                                            Next
                                        End If
                                    End If
                                End If

                                selectSql = "select count(*) as value " &
                                            "from WMS_REPLENISH_D " &
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                                            "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " &
                                            "AND ROD_BATCH_NO = '" & rows.Item("GRA_BATCH_NO").ToString.Trim & "' " &
                                            "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' " &
                                            "AND ROD_PALLET_NO = '" & gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "000") & "' "

                                If DB.getValueFromSQL(selectSql, gConn, transaction) <> "0" Then

                                    updateSql = "update WMS_REPLENISH_D " &
                                                "set ROD_POST_QTY = ISNULL(ROD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) & ", " &
                                                    "sys_lub = '" & Session("usr_id") & "', " &
                                                    "sys_lud = Getdate() " &
                                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                                                "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " &
                                                "AND ROD_BATCH_NO = '" & rows.Item("GRA_BATCH_NO").ToString.Trim & "' " &
                                                "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' " &
                                                "AND ROD_PALLET_NO = '" & gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "000") & "' "

                                    gDB.amendData(updateSql, gConn, transaction)

                                Else

                                    updateSql = "update WMS_REPLENISH_D " &
                                                "set ROD_POST_QTY = ISNULL(ROD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) & ", " &
                                                    "sys_lub = '" & Session("usr_id") & "', " &
                                                    "sys_lud = Getdate() " &
                                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                                                "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " &
                                                "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' "

                                    gDB.amendData(updateSql, gConn, transaction)

                                End If

                            Next
                            'updateSql = "update WMS_REPLENISH " &
                            '            "set RO_STATUS = 'CLOSED', " &
                            '            "sys_lub = '" & Session("usr_id") & "', " &
                            '            "sys_lud = Getdate() " &
                            '            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                            '            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            '            "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                            '            "AND NOT EXISTS (" &
                            '                "SELECT 1 FROM WMS_REPLENISH_D D WHERE ISNULL(ROD_POST_QTY,0) <> ISNULL(ROD_QTY,0) " &
                            '                "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " &
                            '                "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " &
                            '                "AND D.RO_CODE = WMS_REPLENISH.RO_CODE)"

                            updateSql = "update WMS_REPLENISH " &
                                        "set RO_STATUS = 'CLOSED', " &
                                        "sys_lub = '" & Session("usr_id") & "', " &
                                        "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                                        "AND (SELECT Count(D.RO_CODE) FROM WMS_REPLENISH_D D WHERE D.IMP_CODE = WMS_REPLENISH.IMP_CODE AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE AND D.RO_CODE = WMS_REPLENISH.RO_CODE)= (SELECT Count(D.RO_CODE) FROM WMS_REPLENISH_D D WHERE (ISNULL(ROD_POST_QTY,0) >= ISNULL(ROD_QTY,0)) AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE AND D.RO_CODE = WMS_REPLENISH.RO_CODE)"

                            gDB.amendData(updateSql, gConn, transaction)

                            updateSql = "update WMS_REPLENISH " & _
                                        "set RO_STATUS = 'PAR', " & _
                                        "sys_lub = '" & Session("usr_id") & "', " & _
                                        "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                        "AND EXISTS (" & _
                                            "SELECT 1 FROM WMS_REPLENISH_D D WHERE ISNULL(ROD_POST_QTY,0) <> ISNULL(ROD_QTY,0) " & _
                                            "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " & _
                                            "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " & _
                                            "AND D.RO_CODE = WMS_REPLENISH.RO_CODE) "

                            'gDB.amendData(updateSql, gConn, transaction)

                            updateSql = "update WMS_GOODSRCV " & _
                                        "set GR_STATUS = 'POSTED', " & _
                                        "GR_POSTED_DATE = Getdate(), " & _
                                        "GR_POSTED_BY = '" & Session("usr_id") & "', " & _
                                        "sys_lub = '" & Session("usr_id") & "', " & _
                                        "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                            updateSql = "update wms_goodsrcv_PA " &
                                        "set GRA_IPR_LOT_OUTSTANDING = GRA_PA_QTY " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "

                            gDB.amendData(updateSql, gConn, transaction)

                            'Create inspection records
                            insertSql = "insert into WMS_GOODSRCV_INSP (" & _
                                            "IMP_CODE, STORER_CODE, GR_CODE, GRD_SEQ, GRI_SEQ, " & _
                                            "GRI_ITM_CODE, GRI_PACK_KEY, GRI_PALLET_NO, GRI_BATCH_NO, " & _
                                            "GRI_SERIAL_NO_LIST, GRI_INSP_QTY, GRI_PASS_QTY, GRI_REJ_QTY, GRI_REJ_REASON, GRI_PHOTO, GRI_STATUS, " & _
                                            "SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                        "select d.IMP_CODE, d.STORER_CODE, d.GR_CODE, d.GRD_SEQ, row_number() over(partition by D.IMP_CODE, D.STORER_CODE, D.GR_CODE order by d.GRD_SEQ), " & _
                                            "d.GRD_ITM_CODE, d.GRD_PACK_KEY, d.GRD_PALLET_NO, d.GRD_BATCH_NO, " & _
                                            "null, sum(p.GRA_PA_QTY), 0, 0, null, null, 'PENDING', " & _
                                            "'" & Session("usr_id") & "', Getdate(), '" & Session("usr_id") & "', Getdate() " & _
                                        "from WMS_GOODSRCV_D d, WMS_GOODSRCV_PA p " & _
                                        "where d.IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and d.GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' " & _
                                        "and d.GRD_INSP_REQ = 'Y' " & _
                                        "and d.IMP_CODE = p.IMP_CODE " & _
                                        "and d.STORER_CODE = p.STORER_CODE " & _
                                        "and d.GR_CODE = p.GR_CODE " & _
                                        "and d.GRD_ITM_CODE = p.GRA_ITM_CODE " & _
                                        "and d.GRD_PACK_KEY = p.GRA_PACK_KEY " & _
                                        "and isnull(d.GRD_PALLET_NO, '000') = isnull(p.GRA_PALLET_NO, '000') " & _
                                        "and isnull(d.GRD_BATCH_NO, '') = isnull(p.GRA_BATCH_NO, '') " & _
                                        "group by d.IMP_CODE, d.STORER_CODE, d.GR_CODE, d.GRD_SEQ, d.GRD_ITM_CODE, d.GRD_PACK_KEY, d.GRD_PALLET_NO, d.GRD_BATCH_NO "

                            gDB.amendData(insertSql, gConn, transaction)

                            transaction.Commit()

                            'No need as the page will be reload

                            'GR_STATUS.Text = "POSTED"
                            'ar.sec_viewMode = "Y"
                            'btnPost.Visible = False
                            'btnUnPost.Visible = True
                            'ViewState(Page.ClientID & "_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

                            'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                            'selectItemBtn.Attributes("onclick") = "ItemLookUp('" & STORER_CODE.SelectedValue & "');"
                            'cSBBtn.Attributes("onclick") = "checkSB('" & STORER_CODE.SelectedValue & "');"
                            'cSBBtn.Enabled = True

                            'uiFun.displayMsgNew(updtPnlAlert, "1007", "", Session("gLang"))

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
                            'Throw ex

                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Fail to post Put Away! " & ex.Message, Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "上貨指令失敗! " & ex.Message, Session("gLang"))
                            End If
                        End Try


                        If successFlag = True Then
                            reloadPage("Record has been Posted successfully!")

                            'Dim rmtPost As New RemotePost
                            'rmtPost.Url = "GRMain.aspx"
                            'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                            'rmtPost.Add("GR_CODE", gr_code.text)
                            'rmtPost.alertMsg = "Record has been Posted successfully!"
                            'rmtPost.Post()
                        Else

                        End If
                    End If
                End If
            Else
                If Session("gLang") = "E" Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Put away location cannot be empty! Please check put away list.", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", "儲存位置不能空白! 請檢查儲存清單.", Session("gLang"))
                End If
            End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "沒有可供發布的物件!", Session("gLang"))
            End If
        End If

        load_ModalPopupExtender.Hide()
    End Sub

    Protected Sub addROtoGR()
        Dim itmDict As Dictionary(Of String, itmStruct)
        Dim itmKey As String
        Dim keys As Dictionary(Of String, itmStruct).KeyCollection
        Dim keyArray As String()
        Dim pa_dt As DataTable
        Dim plSeq As Integer
        Dim tmpStruct As itmStruct

        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int,GRD_SEQ)) + 1 from WMS_GOODSRCV_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "
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
            Dim RO_dt As DataTable

            Dim ROItemList As String = ""
            Dim ROListarray As String()
            Dim ROSeqListarray As String()
            Dim qtyListarray As String()

            Dim allowGo As Boolean
            allowGo = True

            ROListarray = Split(ROList.Value, ", ")
            ROSeqListarray = Split(ROSeqList.Value, ", ")

            qtyListarray = Split(qtyList.Value, ", ")

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i As Integer = 0 To ROListarray.Count - 1
                    qtyitemDict.Add(ROSeqListarray(i), qtyListarray(i))
                Next
            End If

            If ROListarray.Count = 0 Then
                If ROList.Value <> "" Then
                    ROItemList = ROList.Value & "_000_" & ROSeqList.Value
                    GR_DOC_NO.Value = ROItemList
                    txt_GR_DOC_NO.Text = GR_DOC_NO.Value
                End If
            Else
                For i = 0 To ROListarray.Count - 1
                    ROItemList = gU.appendToList(ROItemList, ROListarray(i) & "_000_" & ROSeqListarray(i))
                    If GR_DOC_NO.Value <> "" Then
                        If GR_DOC_NO.Value <> ROListarray(i) Then
                            allowGo = False
                        End If
                    End If
                    GR_DOC_NO.Value = ROListarray(i)
                    txt_GR_DOC_NO.Text = GR_DOC_NO.Value
                Next
            End If

            If allowGo = False Then
                GR_DOC_NO.Value = ""
                uiFun.displayMsgNew(updtPnlAlert, "", "Your selected item related to differnet RO, please select again!", "")
            Else
                'txt_GR_DOC_TYPE.Text = "RO"
                'GR_DOC_TYPE.Value = "RO"


                SQLString = "SELECT d.ROD_SEQ, d.ROD_DISP_SEQ, d.ROD_PALLET_NO, d.ROD_CARTON_NO, d.ROD_BATCH_NO, m.RO_REM, " &
                    "d.ROD_REF_NO, d.ROD_ITM_PARENT, d.ROD_ITM_CODE, ISNULL(d.ROD_PACK_KEY,1) as ROD_PACK_KEY, d.ROD_ITM_NAME, " &
                    "d.ROD_QTY, d.ROD_UOM, d.ROD_PCS_PER_UOM, d.ROD_TOT_PCS, d.ROD_LENGTH, d.ROD_WIDTH, d.ROD_HEIGHT, " &
                    "d.ROD_KG, d.ROD_CBM, M.RO_REF_NO, d.ROD_VND_CODE, d.rod_series_no, " &
                    "v.AITM_QTY_PER_CTN, i.ITM_SKU_NO, i.ITM_DESC, d.ROD_UOM2, d.ROD_QTY2, " &
                    "d.ROD_ON_BEHALF, d.ROD_INSP_REQ, " &
                    "M.RO_EDI_PO_NO,M.RO_TYPE, isnull(d.ROD_WH_CODE, M.RO_WH_CODE) as ROD_WH_CODE, M.RO_WH_CODE, " &
                    "M.RO_DEST, convert(varchar, M.RO_DATE, " & DDFORMAT & ") as RO_DATE, " &
                    "(d.ROD_QTY - ISNULL(g.grd_rcv_qty, 0)) as AVAIL_QTY, M.RO_TRACK_NO, " &
                    "convert(varchar, d.ROD_EXPIRY_DATE, " & DDFORMAT & ") as ROD_EXPIRY_DATE, " &
                    "convert(varchar,d.ROD_MANU_DATE, " & DDFORMAT & ") as ROD_MANU_DATE " &
                "from  WMS_REPLENISH M " &
                "inner join WMS_REPLENISH_D d " &
                    "on M.IMP_CODE = D.IMP_CODE " &
                    "and M.STORER_CODE = D.STORER_CODE " &
                    "and M.RO_CODE = D.RO_CODE " &
                "left outer join WMS_ITEM i " &
                    "on d.IMP_CODE = i.IMP_CODE " &
                    "and d.STORER_CODE = i.STORER_CODE " &
                    "and d.ROD_ITM_CODE = i.ITM_CODE " &
                    "and d.ROD_PACK_KEY = i.PACK_KEY " &
                "left outer join V_ALT_VEND_ITEM v " &
                    "on d.IMP_CODE = v.IMP_CODE  " &
                    "and d.STORER_CODE = v.STORER_CODE " &
                    "and d.ROD_ITM_CODE = v.ITM_CODE " &
                    "and d.ROD_PACK_KEY = v.PACK_KEY " &
                "left outer join (" &
                    "select gd.grd_itm_code, gd.grd_pack_key, gd.grd_pallet_no, gd.grd_batch_no, sum(gd.grd_rcv_qty) as grd_rcv_qty " &
                        "from WMS_GOODSRCV gh, WMS_GOODSRCV_D gd " &
                        "where gh.imp_code = gd.imp_code " &
                        "and gh.storer_code = gd.storer_code " &
                        "and gh.gr_code = gd.gr_code " &
                        "and gh.imp_code = '" & Session("IMP_CODE") & "' " &
                        "and gh.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and gh.gr_doc_no = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                        "and ISNULL(gh.gr_status, '') <> 'CANCELLED' " &
                        "group by gd.grd_itm_code, gd.grd_pack_key, gd.grd_pallet_no, gd.grd_batch_no) g " &
                    "on d.ROD_ITM_CODE = g.grd_itm_code " &
                    "and d.ROD_PACK_KEY = g.grd_pack_key " &
                    "and ISNULL(d.ROD_BATCH_NO, '') = ISNULL(g.grd_batch_no, '') " &
                    "and ISNULL(d.ROD_PALLET_NO, '000') = ISNULL(g.grd_pallet_no, '000') " &
                "where d.RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                "AND D.RO_CODE + '_000_' + D.ROD_SEQ IN ('" & Replace(ROItemList, ", ", "', '") & "') " &
                "and d.IMP_CODE = '" & Session("IMP_CODE") & "' "


                SQLString = SQLString & " order by d.ROD_ITM_CODE, d.ROD_PACK_KEY, d.ROD_PALLET_NO, d.ROD_BATCH_NO, d.ROD_DISP_SEQ"
                REM **********************
                RO_dt = gDB.getDataTable(SQLString)

                itmDict = New Dictionary(Of String, itmStruct)

                For i As Integer = 0 To RO_dt.Rows.Count - 1
                    If ViewState("n_cur_seq") = "" Then
                        ViewState("n_cur_seq") = next_seq_no
                    Else
                        temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                        ViewState("n_cur_seq") = temp_seq_no.ToString
                    End If

                    Dim itmQty As Integer = 0

                    If qtyitemDict.ContainsKey(RO_dt.Rows(i).Item("ROD_SEQ").ToString) Then
                        itmQty = gU.decodeEmptyCInt(qtyitemDict(RO_dt.Rows(i).Item("ROD_SEQ").ToString), 0)
                    End If

                    'Generate item dict for put away list
                    'If DB.decodeDBNull(RO_dt.Rows(i).Item("AVAIL_QTY"), 0) > 0 Then
                    'If DB.decodeDBNull(RO_dt.Rows(i).Item("AVAIL_QTY"), 0) > 0 Then
                    itmKey = RO_dt.Rows(i).Item("ROD_ITM_CODE").ToString & "#_#" & _
                            RO_dt.Rows(i).Item("ROD_PACK_KEY").ToString & "#_#" & _
                            RO_dt.Rows(i).Item("ROD_PALLET_NO").ToString & "#_#" & _
                            RO_dt.Rows(i).Item("ROD_BATCH_NO").ToString

                    'If Not itmDict.ContainsKey(itmKey) Then
                    '    itmDict.Add(itmKey, RO_dt.Rows(i).Item("AVAIL_QTY"))
                    'Else
                    '    itmDict.Item(itmKey) = itmDict.Item(itmKey) + RO_dt.Rows(i).Item("AVAIL_QTY")
                    'End If

                    If Not itmDict.ContainsKey(itmKey) Then
                        'itmDict.Add(itmKey, RO_dt.Rows(i).Item("AVAIL_QTY"))
                        tmpStruct = New itmStruct
                        tmpStruct.itmName = RO_dt.Rows(i).Item("ROD_ITM_NAME").ToString

                        'tmpStruct.qty = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AVAIL_QTY"), 0)
                        tmpStruct.qty = itmQty

                        tmpStruct.itmSeq = ViewState("n_cur_seq")

                        tmpStruct.batchNo = RO_dt.Rows(i).Item("ROD_BATCH_NO").ToString
                        tmpStruct.refNo = RO_dt.Rows(i).Item("ROD_REF_NO").ToString

                        tmpStruct.skuNo = RO_dt.Rows(i).Item("ITM_SKU_NO").ToString

                        tmpStruct.expDate = RO_dt.Rows(i).Item("ROD_EXPIRY_DATE").ToString
                        tmpStruct.manuDate = RO_dt.Rows(i).Item("ROD_MANU_DATE").ToString

                        itmDict.Add(itmKey, tmpStruct)
                    Else
                        'itmDict.Item(itmKey) = itmDict.Item(itmKey) + RO_dt.Rows(i).Item("AVAIL_QTY")
                        tmpStruct = itmDict.Item(itmKey)
                        'tmpStruct.qty = itmDict.Item(itmKey).qty + gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AVAIL_QTY"), 0)
                        tmpStruct.qty = itmDict.Item(itmKey).qty + itmQty
                    End If
                    'End If

                    dt.Rows.Add()

                    rows_count = dt.Rows.Count



                    REM **********************
                    REM Modify Here
                    dt.Rows(rows_count - 1).Item("grd_seq") = ViewState("n_cur_seq").ToString
                    dt.Rows(rows_count - 1).Item("GRD_DISP_SEQ") = ViewState("n_cur_seq").ToString 'RO_dt.Rows(i).Item("ROD_DISP_SEQ")
                    dt.Rows(rows_count - 1).Item("GRD_ITM_CODE") = RO_dt.Rows(i).Item("ROD_ITM_CODE")
                    dt.Rows(rows_count - 1).Item("GRD_ITM_NAME") = RO_dt.Rows(i).Item("ROD_ITM_NAME")
                    dt.Rows(rows_count - 1).Item("GRD_PACK") = ""
                    dt.Rows(rows_count - 1).Item("GRD_PACK_NO") = ""
                    If CDbl(RO_dt.Rows(i).Item("AVAIL_QTY")) < 0 Then
                        dt.Rows(rows_count - 1).Item("GRD_PO_QTY") = 0
                        'dt.Rows(rows_count - 1).Item("GRD_RCV_QTY") = 0
                        'dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = 0
                    Else
                        dt.Rows(rows_count - 1).Item("GRD_PO_QTY") = RO_dt.Rows(i).Item("AVAIL_QTY")
                        'dt.Rows(rows_count - 1).Item("GRD_RCV_QTY") = RO_dt.Rows(i).Item("AVAIL_QTY")
                        'dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AVAIL_QTY"), 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_PCS_PER_UOM"), 0)
                    End If

                    dt.Rows(rows_count - 1).Item("GRD_RCV_QTY") = itmQty
                    dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = itmQty * If(RO_dt.Rows(i).Item("ROD_PCS_PER_UOM") IsNot DBNull.Value, RO_dt.Rows(i).Item("ROD_PCS_PER_UOM").ToString(), "0")
                    'dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = itmQty * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_PCS_PER_UOM"), 0)

                    dt.Rows(rows_count - 1).Item("GRD_OS_QTY") = CInt(dt.Rows(rows_count - 1).Item("GRD_PO_QTY")) - itmQty

                    'dt.Rows(rows_count - 1).Item("GRD_WH") = ""
                    'dt.Rows(rows_count - 1).Item("GRD_LOC") = ""
                    dt.Rows(rows_count - 1).Item("GRD_PALLET_NO") = RO_dt.Rows(i).Item("ROD_PALLET_NO")
                    dt.Rows(rows_count - 1).Item("GRD_CARTON_NO") = RO_dt.Rows(i).Item("ROD_CARTON_NO")
                    dt.Rows(rows_count - 1).Item("GRD_BATCH_NO") = RO_dt.Rows(i).Item("ROD_BATCH_NO")
                    dt.Rows(rows_count - 1).Item("GRD_REF_NO") = RO_dt.Rows(i).Item("ROD_REF_NO")
                    'dt.Rows(rows_count - 1).Item("GRD_BRAND") = ""
                    'dt.Rows(rows_count - 1).Item("GRD_SERIES") = RO_dt.Rows(i).Item("ROD_SERIES_NO")
                    'dt.Rows(rows_count - 1).Item("GRD_MODEL") = ""
                    dt.Rows(rows_count - 1).Item("GRD_UOM") = RO_dt.Rows(i).Item("ROD_UOM")
                    dt.Rows(rows_count - 1).Item("GRD_LENGTH") = RO_dt.Rows(i).Item("ROD_LENGTH")
                    dt.Rows(rows_count - 1).Item("GRD_WIDTH") = RO_dt.Rows(i).Item("ROD_WIDTH")
                    dt.Rows(rows_count - 1).Item("GRD_HEIGHT") = RO_dt.Rows(i).Item("ROD_HEIGHT")

                    dt.Rows(rows_count - 1).Item("GRD_DESTINATION") = RO_dt.Rows(i).Item("RO_DEST")
                    dt.Rows(rows_count - 1).Item("GRD_STOCKTYPE") = ""

                    dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = RO_dt.Rows(i).Item("ITM_SKU_NO")
                    dt.Rows(rows_count - 1).Item("GRD_PACK_KEY") = RO_dt.Rows(i).Item("ROD_PACK_KEY")

                    'dt.Rows(rows_count - 1).Item("ITM_DESC") = RO_dt.Rows(i).Item("ITM_DESC")

                    dt.Rows(rows_count - 1).Item("GRD_PCS_PER_UOM") = RO_dt.Rows(i).Item("ROD_PCS_PER_UOM")
                    dt.Rows(rows_count - 1).Item("GRD_VND_CODE") = RO_dt.Rows(i).Item("ROD_VND_CODE")

                    dt.Rows(rows_count - 1).Item("GRD_EXPIRY_DATE") = RO_dt.Rows(i).Item("ROD_EXPIRY_DATE")
                    dt.Rows(rows_count - 1).Item("GRD_MANU_DATE") = RO_dt.Rows(i).Item("ROD_MANU_DATE")

                    dt.Rows(rows_count - 1).Item("GRD_WH") = RO_dt.Rows(i).Item("ROD_WH_CODE")

                    dt.Rows(rows_count - 1).Item("GRD_UOM2") = RO_dt.Rows(i).Item("ROD_UOM2")
                    dt.Rows(rows_count - 1).Item("GRD_QTY2") = RO_dt.Rows(i).Item("ROD_QTY2")

                    'dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = RO_dt.Rows(i).Item("ROD_TOT_PCS")

                    dt.Rows(rows_count - 1).Item("GRD_PCS_PER_CARTON") = RO_dt.Rows(i).Item("AITM_QTY_PER_CTN")

                    If RO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString.Trim <> "" AndAlso RO_dt.Rows(i).Item("AITM_QTY_PER_CTN").ToString.Trim <> "0" Then
                        dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON") = Math.Ceiling(itmQty / CDbl(RO_dt.Rows(i).Item("AITM_QTY_PER_CTN")))
                    Else
                        dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON") = 0
                    End If

                    'If gU.decodeNullOrEmpty(dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON").ToString.Trim, "0") <> "0" Then
                    '    dt.Rows(rows_count - 1).Item("GRD_KG") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_KG").ToString.Trim, 0) * gU.decodeEmptyCdbl(dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON").ToString.Trim, 0)
                    'Else
                    '    dt.Rows(rows_count - 1).Item("GRD_KG") = RO_dt.Rows(i).Item("ROD_KG")
                    'End If
                    dt.Rows(rows_count - 1).Item("GRD_KG") = RO_dt.Rows(i).Item("ROD_KG")


                    'dt.Rows(rows_count - 1).Item("GRD_ON_BEHALF") = RO_dt.Rows(i).Item("ROD_ON_BEHALF")

                    'If dt.Rows(rows_count - 1).Item("GRD_ON_BEHALF").ToString.Trim = "Y" Then
                    '    dt.Rows(rows_count - 1).Item("GRD_STATUS") = "WAITCOLLECT"
                    'Else
                    '    dt.Rows(rows_count - 1).Item("GRD_STATUS") = "NEW"
                    'End If

                    dt.Rows(rows_count - 1).Item("GRD_INSP_REQ") = RO_dt.Rows(i).Item("ROD_INSP_REQ")

                    'If gU.decodeNullOrEmpty(dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON").ToString.Trim, "0") <> "0" Then
                    '    dt.Rows(rows_count - 1).Item("GRD_CBM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_HEIGHT").ToString.Trim, 0) / 1000000 * gU.decodeEmptyCdbl(dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON").ToString.Trim, 0)
                    '    'dt.Rows(rows_count - 1).Item("GRD_CBM") = CDbl(RO_dt.Rows(i).Item("ROD_CBM")) * dt.Rows(rows_count - 1).Item("GRD_NO_OF_CARTON")
                    'Else
                    '    dt.Rows(rows_count - 1).Item("GRD_CBM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_HEIGHT").ToString.Trim, 0) / 1000000
                    '    'dt.Rows(rows_count - 1).Item("GRD_CBM") = CDbl(RO_dt.Rows(i).Item("ROD_CBM"))
                    'End If

                    'New CBM logic by qty
                    If itmQty > 0 Then
                        dt.Rows(rows_count - 1).Item("GRD_CBM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ROD_HEIGHT").ToString.Trim, 0) / 1000000 * itmQty
                    Else
                        dt.Rows(rows_count - 1).Item("GRD_CBM") = 0
                    End If


                    'PO_DATE.Text = cU.chgToYYYYMMDD(RO_dt.Rows(i).Item("RO_DATE").ToString)
                    PO_DATE.Text = RO_dt.Rows(i).Item("RO_DATE").ToString
                    GR_TRACK_NO.Value = RO_dt.Rows(i).Item("RO_TRACK_NO").ToString



                    REM **********************
                    dt.Rows(rows_count - 1).Item("mFlag") = "N"

                    GR_REF_NO.Value = gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("RO_REF_NO").ToString, "")

                    GR_EDI_PO_NO.Text = gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("RO_EDI_PO_NO").ToString, "")

                    GR_REM.Text = gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("RO_REM").ToString, "")

                    GR_WH_CODE.Text = RO_dt.Rows(i).Item("RO_WH_CODE").ToString.Trim

                    txt_GR_DOC_TYPE.Text = gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("RO_TYPE").ToString, "")
                    GR_DOC_TYPE.Value = gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("RO_TYPE").ToString, "")
                Next
                dt.AcceptChanges()
                Session("dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()

                'Generate put awya list
                'If itmDict.Count > 0 Then

                '    Dim paTable As PutAwayTable

                '    paTable = New PutAwayTable(dt, Session("_M_IB_GR_TMP_pa_dt"), Session("_M_IB_GR_TMP_pa_seq"), Session("_M_IB_GR_TMP_pa_del_list"))

                '    paTable.genPutAway(True)

                '    Session("_M_IB_GR_TMP_pa_del_list") = paTable.putAwayDelList

                '    Session("_M_IB_GR_TMP_pa_seq") = paTable.putAwaySeq

                'End If
            End If
        End If
    End Sub

    Protected Sub addItemtoGR()
        Dim itmDict As Dictionary(Of String, itmStruct)
        Dim itmKey As String
        Dim keys As Dictionary(Of String, itmStruct).KeyCollection
        Dim keyArray As String()
        Dim pa_dt As DataTable
        Dim plSeq As Integer
        Dim tmpStruct As itmStruct

        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int,GRD_SEQ)) + 1 from WMS_GOODSRCV_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "
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
            Dim GR_dt As DataTable

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


            'SQLString = "SELECT M.*, D.* " & _
            '"from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
            '"where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            '"and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
            '"AND M.ITM_CODE + '_000_' + M.PACK_KEY IN ('" & Replace(itemPackList, ", ", "', '") & "') " & _
            '"AND M.IMP_CODE *= D.IMP_CODE " & _
            '"AND M.STORER_CODE *= D.STORER_CODE " & _
            '"AND M.PACK_KEY *= D.PACK_KEY " & _
            '"AND M.ITM_CODE *= D.ITM_CODE"

            'SQLString = SQLString & " order by M.ITM_CODE, M.PACK_KEY, VND_CODE"

            SQLString = ""
            SQLString += "select wms_item.*, wms_alt_vend_item.* "
            SQLString += "from wms_item left outer join wms_alt_vend_item on "
            SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_code + '_000_' + wms_item.pack_key + '_000_' + ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

            SQLString = SQLString & " order by wms_item.itm_code, wms_item.pack_key, wms_alt_vend_item.vnd_code"
            REM **********************
            GR_dt = gDB.getDataTable(SQLString)

            itmDict = New Dictionary(Of String, itmStruct)

            For i As Integer = 0 To GR_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                'Generate item dict for put away list
                If DB.decodeDBNull(GR_dt.Rows(i).Item("ITM_BALANCE"), 0) > 0 Then
                    itmKey = GR_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & GR_dt.Rows(i).Item("PACK_KEY").ToString

                    'If Not itmDict.ContainsKey(itmKey) Then
                    '    itmDict.Add(itmKey, GR_dt.Rows(i).Item("ITM_BALANCE"))
                    'Else
                    '    itmDict.Item(itmKey) = itmDict.Item(itmKey) + GR_dt.Rows(i).Item("ITM_BALANCE")
                    'End If

                    If Not itmDict.ContainsKey(itmKey) Then
                        'itmDict.Add(itmKey, GR_dt.Rows(i).Item("ITM_BALANCE"))
                        tmpStruct = New itmStruct
                        tmpStruct.itmName = GR_dt.Rows(i).Item("ITM_NAME")
                        tmpStruct.qty = GR_dt.Rows(i).Item("ITM_BALANCE")
                        tmpStruct.itmSeq = ViewState("n_cur_seq")
                        itmDict.Add(itmKey, tmpStruct)
                    Else
                        'itmDict.Item(itmKey) = itmDict.Item(itmKey) + GR_dt.Rows(i).Item("ITM_BALANCE")
                        tmpStruct = itmDict.Item(itmKey)
                        tmpStruct.qty = tmpStruct.qty + GR_dt.Rows(i).Item("ITM_BALANCE")
                    End If
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("grd_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("GRD_DISP_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("GRD_ITM_CODE") = GR_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("GRD_ITM_NAME") = GR_dt.Rows(i).Item("ITM_NAME")

                If gU.decodeEmptyCdbl(GR_dt.Rows(i).Item("ITM_BALANCE").ToString, 0) < 0 Then
                    dt.Rows(rows_count - 1).Item("GRD_PO_QTY") = 0
                    dt.Rows(rows_count - 1).Item("GRD_RCV_QTY") = 0
                Else
                    dt.Rows(rows_count - 1).Item("GRD_PO_QTY") = GR_dt.Rows(i).Item("ITM_BALANCE")
                    dt.Rows(rows_count - 1).Item("GRD_RCV_QTY") = GR_dt.Rows(i).Item("ITM_BALANCE")
                End If


                dt.Rows(rows_count - 1).Item("GRD_PACK_KEY") = GR_dt.Rows(i).Item("PACK_KEY")

                dt.Rows(rows_count - 1).Item("GRD_OS_QTY") = 0
                dt.Rows(rows_count - 1).Item("GRD_SERIES") = ""
                dt.Rows(rows_count - 1).Item("GRD_UOM") = GR_dt.Rows(i).Item("AITM_UOM")

                dt.Rows(rows_count - 1).Item("GRD_KG") = GR_dt.Rows(i).Item("ITM_BALANCE_KG")
                dt.Rows(rows_count - 1).Item("GRD_CBM") = GR_dt.Rows(i).Item("ITM_BALANCE_CBM")
                dt.Rows(rows_count - 1).Item("GRD_PACK_KEY") = GR_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("GRD_PCS_PER_UOM") = "1"
                dt.Rows(rows_count - 1).Item("GRD_TOT_PCS") = 0
                dt.Rows(rows_count - 1).Item("GRD_VND_CODE") = GR_dt.Rows(i).Item("VND_CODE")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
            Next
            dt.AcceptChanges()
            Session("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

            ar.hideGVForStorer(GridView1, STORER_CODE.Text, "GR", "WMS_GOODSRCV_D")

            'Generate put awya list
            If itmDict.Count > 0 Then
                keys = itmDict.Keys
                pa_dt = Session("_M_IB_GR_TMP_pa_dt")

                plSeq = CInt(Session("_M_IB_GR_TMP_pa_seq"))

                For i = 0 To keys.Count - 1
                    pa_dt.Rows.Add()

                    plSeq = plSeq + 1

                    rows_count = pa_dt.Rows.Count

                    keyArray = Split(keys(i), "#_#")

                    pa_dt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
                    pa_dt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
                    pa_dt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
                    pa_dt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
                    pa_dt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
                    'pa_dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = itmDict.Item(keys(i)).qty
                    pa_dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
                    pa_dt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
                    pa_dt.Rows(rows_count - 1).Item("GRA_WH") = ""
                    pa_dt.Rows(rows_count - 1).Item("GRA_LOC") = ""
                    pa_dt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = ""
                    pa_dt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
                    pa_dt.Rows(rows_count - 1).Item("mFlag") = "N"
                Next

                Session("_M_IB_GR_TMP_pa_seq") = CStr(plSeq)

                pa_dt.AcceptChanges()
            End If
        End If

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim valL As String = TryCast(GridView1.Rows(e.RowIndex).FindControl("grd_length"), TextBox).Text
        Dim valW As String = TryCast(GridView1.Rows(e.RowIndex).FindControl("grd_width"), TextBox).Text
        Dim valH As String = TryCast(GridView1.Rows(e.RowIndex).FindControl("grd_height"), TextBox).Text
        Dim valKG As String = TryCast(GridView1.Rows(e.RowIndex).FindControl("grd_kg"), TextBox).Text

        Dim putValL, putValW, putValH, putValKG As Double

        If IsNumeric(valL) Then
            putValL = Convert.ToDouble(valL)
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Length!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小 - 長度!", Session("gLang"))
            End If
        End If

        If IsNumeric(valW) Then
            putValW = Convert.ToDouble(valW)
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Width!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小 - 闊度!", Session("gLang"))
            End If
        End If

        If IsNumeric(valH) Then
            putValH = Convert.ToDouble(valH)
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Size Height!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 大小 - 高度!", Session("gLang"))
            End If
        End If

        If IsNumeric(valKG) Then
            putValKG = Convert.ToDouble(valKG)
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Weight(kg)!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 重量(kg)!", Session("gLang"))
            End If
        End If

        ViewState("COPYFROM_L") = FormatNumber(putValL, 2).ToString
        ViewState("COPYFROM_W") = FormatNumber(putValW, 2).ToString
        ViewState("COPYFROM_H") = FormatNumber(putValH, 2).ToString
        ViewState("COPYFROM_KG") = FormatNumber(putValKG, 2).ToString

        DirectCast(GridView1.Rows(e.RowIndex).FindControl("grd_length"), TextBox).Text = FormatNumber(putValL, 2).ToString
        DirectCast(GridView1.Rows(e.RowIndex).FindControl("grd_width"), TextBox).Text = FormatNumber(putValW, 2).ToString
        DirectCast(GridView1.Rows(e.RowIndex).FindControl("grd_height"), TextBox).Text = FormatNumber(putValH, 2).ToString
        DirectCast(GridView1.Rows(e.RowIndex).FindControl("grd_kg"), TextBox).Text = FormatNumber(putValKG, 2).ToString

        ViewState("ISCOPYING") = "Y"

        'If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
        '    Me.ClientScript.RegisterStartupScript(Me.GetType, "iscopied", "alert('Size(LxWxH): " & FormatNumber(putValL, 2).ToString & _
        '                                                                        "x" & FormatNumber(putValW, 2).ToString & _
        '                                                                        "x" & FormatNumber(putValH, 2).ToString & _
        '                                                                        "cm, Weight: " & FormatNumber(putValKG, 2).ToString & "kg has been copied.')", True)
        'End If
    End Sub

    Private Sub updtBatchNo()
        Dim i As Integer
        Dim newBatchNo As String

        If ViewState("STO_BATCH_FIELD_REF") <> "" Then
            If GridView1.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    newBatchNo = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), CType(GridView1.Rows(i).FindControl("GRD_BATCH_NO"), TextBox).Text, CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim, CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim)

                    If newBatchNo <> CType(GridView1.Rows(i).FindControl("GRD_BATCH_NO"), TextBox).Text Then

                        'uiFun.load_ComboBox(CType(GridView1.Rows(i).FindControl("GRD_BATCH_NO"), AjaxControlToolkit.ComboBox),
                        '            "select '" & newBatchNo & "' as dc_date_code  union select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", , False)

                        CType(GridView1.Rows(i).FindControl("GRD_BATCH_NO"), TextBox).Text = newBatchNo

                    End If
                Next
            End If
        End If
    End Sub

    Private Sub updtPutAwayPallet()
        Dim paTable As PutAwayTable
        Dim pa_dt As DataTable
        'Dim itmDict As Dictionary(Of String, Double)
        Dim i, j As Integer
        Dim chkUpdt, genUpdtPa As Boolean
        'Dim plSeq As Integer
        Dim rows_count As Integer = 0

        'paDt = Session("_M_IB_GR_TMP_pa_dt")

        If GridView1.Rows.Count > 0 Then
            pa_dt = Session("_M_IB_GR_TMP_pa_dt")

            'itmDict = New Dictionary(Of String, Double)

            If pa_dt Is Nothing OrElse pa_dt.Rows.Count = 0 Then
                'paTable = New PutAwayTable(dt, Session("_M_IB_GR_TMP_pa_dt"), Session("_M_IB_GR_TMP_pa_seq"), Session("_M_IB_GR_TMP_pa_del_list"))

                'paTable.genPutAway(True)

                'Session("_M_IB_GR_TMP_pa_del_list") = paTable.putAwayDelList

                'Session("_M_IB_GR_TMP_pa_seq") = paTable.putAwaySeq

            Else
                genUpdtPa = False

                For i = 0 To GridView1.Rows.Count - 1
                    If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text, "0") > 0 Then
                        chkUpdt = False

                        For j = 0 To pa_dt.Rows.Count - 1
                            If pa_dt.Rows(j).Item("GRA_ITM_CODE").ToString = CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text.Trim AndAlso _
                                pa_dt.Rows(j).Item("GRA_PACK_KEY").ToString = CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text.Trim Then

                                If pa_dt.Rows(j).Item("GRA_SPLIT_FR").ToString = CType(GridView1.Rows(i).FindControl("grd_seq"), HiddenField).Value.Trim Then
                                    If pa_dt.Rows(j).Item("GRA_PALLET_NO").ToString <> CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim Then
                                        pa_dt.Rows(j).Item("GRA_PALLET_NO") = CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim
                                    End If

                                    If pa_dt.Rows(j).Item("gra_ref_no").ToString <> CType(GridView1.Rows(i).FindControl("grd_ref_no"), TextBox).Text.Trim Then
                                        pa_dt.Rows(j).Item("gra_ref_no") = CType(GridView1.Rows(i).FindControl("grd_ref_no"), TextBox).Text.Trim
                                    End If

                                    If pa_dt.Rows(j).Item("GRA_BATCH_NO").ToString <> CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text.Trim Then
                                        pa_dt.Rows(j).Item("GRA_BATCH_NO") = CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text.Trim
                                    End If

                                    If pa_dt.Rows(j).Item("GRA_EXPIRY_DATE").ToString <> CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim Then
                                        pa_dt.Rows(j).Item("GRA_EXPIRY_DATE") = CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim
                                    End If

                                    If pa_dt.Rows(j).Item("GRA_MANU_DATE").ToString <> CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim Then
                                        pa_dt.Rows(j).Item("GRA_MANU_DATE") = CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim
                                    End If

                                    chkUpdt = True
                                ElseIf pa_dt.Rows(j).Item("GRA_PALLET_NO").ToString = CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim AndAlso
                                    pa_dt.Rows(j).Item("GRA_BATCH_NO").ToString = CType(GridView1.Rows(i).FindControl("grd_batch_no"), TextBox).Text.Trim Then

                                    chkUpdt = True
                                End If
                            End If
                        Next

                        If Not chkUpdt Then
                            genUpdtPa = True
                        End If

                        'If Not chkUpdt Then
                        '    plSeq = CInt(Session("_M_IB_GR_TMP_pa_seq"))

                        '    pa_dt.Rows.Add()

                        '    plSeq = plSeq + 1

                        '    rows_count = pa_dt.Rows.Count

                        '    pa_dt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = CType(GridView1.Rows(i).FindControl("grd_itm_code"), Label).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = CType(GridView1.Rows(i).FindControl("grd_itm_name"), Label).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = CType(GridView1.Rows(i).FindControl("itm_sku_no"), Label).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = CType(GridView1.Rows(i).FindControl("grd_pack_key"), Label).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = CType(GridView1.Rows(i).FindControl("grd_pallet_no"), TextBox).Text.Trim
                        '    'pa_dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_SUG_QTY") = gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("grd_rcv_qty"), TextBox).Text, "0")
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = CType(GridView1.Rows(i).FindControl("grd_seq"), HiddenField).Value.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = CType(GridView1.Rows(i).FindControl("GRD_BATCH_NO"), AjaxControlToolkit.ComboBox).SelectedValue

                        '    pa_dt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = CType(GridView1.Rows(i).FindControl("GRD_EXPIRY_DATE"), TextBox).Text.Trim
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = CType(GridView1.Rows(i).FindControl("GRD_MANU_DATE"), TextBox).Text.Trim

                        '    pa_dt.Rows(rows_count - 1).Item("GRA_WH") = ""
                        '    pa_dt.Rows(rows_count - 1).Item("GRA_LOC") = ""
                        '    pa_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                        '    Session("_M_IB_GR_TMP_pa_seq") = CStr(plSeq)
                        'End If
                    End If
                Next

                If genUpdtPa Then
                    paTable = New PutAwayTable(dt, Session("_M_IB_GR_TMP_pa_dt"), Session("_M_IB_GR_TMP_pa_seq"), Session("_M_IB_GR_TMP_pa_del_list"))

                    paTable.genPutAway(False)

                    Session("_M_IB_GR_TMP_pa_del_list") = paTable.putAwayDelList

                    Session("_M_IB_GR_TMP_pa_seq") = paTable.putAwaySeq
                End If
            End If



            'pa_dt.AcceptChanges()
        End If
    End Sub

    Protected Sub btnPutAway_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPutAway.Click

        If editMode.Value <> "V" Then

            updtBatchNo()

            If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                Session("dt") = dt

                updtPutAwayPallet()
            End If

            Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
            updtPnlLinkBar.Update()

            GridView1.DataSource = dt
            GridView1.DataBind()

            'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            updtPnlItemGV.Update()
        End If
        'isClickPutAway = True

        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "SHOW_PUTAWAY", "putAwayMain();", True)
    End Sub

    Private Sub reOrderGRD(ByRef impDt As DataTable)
        Dim tempDT As New DataTable

        tempDT = impDt.Clone

        Dim dispSeq As Integer = 1

        Dim rowsFound = From c In impDt.Rows _
                          Order By _
                          Convert.ToInt32(c.item("grd_disp_seq")) _
                          Ascending _
                         Select c

        Dim sortRows As Object()

        sortRows = rowsFound.ToArray

        For i As Integer = 0 To sortRows.Count - 1
            If sortRows(i).item("grd_ref_seq").ToString = "" Then
                sortRows(i).Item("grd_disp_seq") = dispSeq

                sortRows(i).AcceptChanges()

                dispSeq += 1

                tempDT.ImportRow(sortRows(i))

                If sortRows(i).Item("mFlag").ToString = "D" Then
                    Dim splitRows As DataRow() = impDt.Select("grd_ref_seq = '" & sortRows(i).item("grd_seq").ToString & "'")

                    For Each r As DataRow In splitRows
                        r.Item("mFlag") = "D"
                        r.AcceptChanges()
                    Next
                End If

            ElseIf sortRows(i).item("grd_ref_seq").ToString <> "" Then
                Dim refRows As DataRow() = impDt.Select("grd_seq = '" & sortRows(i).item("grd_ref_seq").ToString & "'")

                If refRows.Count > 0 Then
                    sortRows(i).Item("grd_disp_seq") = refRows(0).Item("grd_disp_seq")
                    sortRows(i).AcceptChanges()

                    tempDT.ImportRow(sortRows(i))
                End If
            End If
        Next

        tempDT.AcceptChanges()

        impDt = tempDT
    End Sub

    Protected Sub btnUnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUnPost.Click
        Dim selectSql, updateSql, deleteSql As String
        Dim gConn As SqlConnection
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim itmKey As String
        Dim pa_dt, snDt As DataTable
        Dim SQLString1 As String
        Dim checkdt As DataTable
        Dim holddt As DataTable
        Dim coHolddt As DataTable
        Dim SQLString As String
        Dim successFlag As Boolean
        Dim dmgLoc As String
        Dim errorMsg As String = ""

        SQLString1 = "SELECT GR_STATUS FROM wms_goodsrcv M " & _
            "WHERE GR_STATUS <> 'POSTED' " & _
            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString1)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The GR status is New, please post before un-post GR!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The GR status is New, please post before un-post GR!", Session("gLang"))
            End If
            Exit Sub
        End If

        If GridView1.Rows.Count > 0 Then
            If GR_STATUS.Text = "POSTED" Then
                gConn = gDB.getConnection()

                Dim transaction As SqlTransaction
                transaction = gConn.BeginTransaction()

                Try
                    SQLString = "select min(loc) as value " & _
                                "from v_location " & _
                                "where ar_damage_yn = 'Y' " & _
                                "and isnull(ar_insp_area, 'N') <> 'Y' " & _
                                "and wh_code = '" & gU.dbEncode(GR_WH_CODE.Text.Trim) & "' "

                    dmgLoc = DB.getValueFromSQL(SQLString, gConn, transaction)

                    If dmgLoc = "" Then
                        SQLString = "select min(loc) as value " & _
                                    "from v_location " & _
                                    "where ar_damage_yn = 'Y' " & _
                                    "and isnull(ar_insp_area, 'N') <> 'Y' " & _
                                    "and wh_code = '" & gU.dbEncode(GR_WH_CODE.Text.Trim) & "' "

                        dmgLoc = DB.getValueFromSQL(SQLString, gConn, transaction)
                    End If

                    qtyDict = New Dictionary(Of String, Double)
                    cbmDict = New Dictionary(Of String, Double)
                    wgtDict = New Dictionary(Of String, Double)

                    For Each rows As DataRow In dt.Rows
                        Dim cbm As Double = gU.decodeEmptyCdbl(rows.Item("grd_length").ToString, 0) * _
                                                gU.decodeEmptyCdbl(rows.Item("grd_width").ToString, 0) * _
                                                gU.decodeEmptyCdbl(rows.Item("grd_height").ToString, 0)

                        If cbm <> 0 Then cbm = cbm / 1000000


                        itmKey = gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, "") & "#_#" & _
                                gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, "") & "#_#" & _
                                gU.decodeNull(rows.Item("grd_pallet_no").ToString.Trim, "") & "#_#" & _
                                gU.decodeNull(rows.Item("grd_batch_no").ToString.Trim, "")

                        If qtyDict.ContainsKey(itmKey) Then
                            qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("grd_rcv_qty").ToString.Trim, 0)
                            cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(cbm, 0), 14)
                            wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("GRD_KG").ToString.Trim, 0), 14)
                        Else
                            qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("grd_rcv_qty").ToString.Trim, 0))
                            cbmDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(cbm, 0), 14))
                            wgtDict.Add(itmKey, Math.Round(gU.decodeEmptyCdbl(rows.Item("GRD_KG").ToString.Trim, 0), 14))
                        End If
                    Next

                    pa_dt = Session("_M_IB_GR_TMP_pa_dt")

                    snDt = Session("sn_dt")

                    REM Hold Logic
                    SQLString = "select gh.gr_code, grd_itm_code, grd_batch_no, grd_pack_key, grd_pallet_no, grd_rcv_qty " & _
                                "from wms_goodsrcv gh, wms_goodsrcv_d gd " & _
                                "where gh.storer_code = gd.storer_code " & _
                                "and gh.imp_code = gd.imp_code  " & _
                                "and gh.gr_code = gd.gr_code " & _
                                "and gh.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                "and gh.imp_code = '" & Session("imp_code") & "' " & _
                                "and gh.gr_code = '" & gU.dbEncode(GR_CODE.Text) & "' "
                    SQLString = SQLString & " order by 1, 2"

                    holddt = gDB.getDataTable(SQLString)
                    For i As Integer = 0 To holddt.Rows.Count - 1
                        Dim grd_rcv_qty As Integer = 0
                        grd_rcv_qty = gU.decodeEmptyCInt(holddt.Rows(i).Item("grd_rcv_qty"), 0)
                        SQLString = "select hold.co_code, hold.cod_seq, hold.coh_seq, COH_IN_STOCK_QTY, COH_QTY " & _
                                "from wms_cust_order_hold hold, wms_cust_order_d cd " & _
                                "where hold.storer_code = cd.storer_code " & _
                                "and hold.imp_code = cd.imp_code " & _
                                "and hold.co_code = cd.co_code " & _
                                "and hold.cod_seq = cd.cod_seq " & _
                                "and hold.coh_status = 'HOLD' " & _
                                "and hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                "and hold.ro_code = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                "and hold.coh_itm_code = '" & gU.dbEncode(holddt.Rows(i).Item("grd_itm_code").ToString) & "' " & _
                                "and ISNULL(hold.coh_batch_no, '') = ISNULL('" & gU.dbEncode(holddt.Rows(i).Item("grd_batch_no").ToString) & "', '') " & _
                                "and hold.coh_pack_key = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pack_key").ToString) & "' " & _
                                "and isnull(hold.coh_pallet_no, '000') = isnull('" & gU.dbEncode(holddt.Rows(i).Item("grd_pallet_no").ToString) & "', '000') "

                        '"and cd.cod_itm_code = '" & gU.dbEncode(holddt.Rows(i).Item("grd_itm_code").ToString) & "' " & _
                        '        "and ISNULL(cd.cod_batch_no, '') = ISNULL('" & gU.dbEncode(holddt.Rows(i).Item("grd_batch_no").ToString) & "', '') " & _
                        '        "and cd.cod_pack_key = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pack_key").ToString) & "' " & _
                        '        "and cd.cod_pallet_no = '" & gU.dbEncode(holddt.Rows(i).Item("grd_pallet_no").ToString) & "' "

                        SQLString = SQLString & " order by 1, 2, 3"

                        coHolddt = gDB.getDataTable(SQLString)
                        For j As Integer = 0 To coHolddt.Rows.Count - 1
                            Dim COH_IN_STOCK_QTY As Integer = 0
                            If gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0) - grd_rcv_qty > 0 Then
                                COH_IN_STOCK_QTY = 0
                                grd_rcv_qty = grd_rcv_qty - gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0)
                            Else
                                COH_IN_STOCK_QTY = gU.decodeEmptyCInt(coHolddt.Rows(j).Item("COH_IN_STOCK_QTY"), 0) - grd_rcv_qty
                                grd_rcv_qty = 0
                            End If
                            updateSql = "UPDATE wms_cust_order_hold hold SET COH_IN_STOCK_QTY = " & COH_IN_STOCK_QTY & " " & _
                                        "WHERE hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                        "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                        "and hold.co_code = '" & gU.dbEncode(coHolddt.Rows(j).Item("co_code").ToString) & "' " & _
                                        "and hold.cod_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("cod_seq").ToString) & "' " & _
                                        "and hold.coh_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("coh_seq").ToString) & "' " & _
                                        "and hold.coh_status = 'HOLD' "
                            gDB.amendData(updateSql, gConn, transaction)

                            REM update to hold Stock
                            updateSql = "UPDATE wms_cust_order_hold hold SET coh_type = 'RO' " & _
                                        "WHERE hold.storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                        "and hold.imp_code = '" & Session("imp_code") & "' " & _
                                        "and hold.co_code = '" & gU.dbEncode(coHolddt.Rows(j).Item("co_code").ToString) & "' " & _
                                        "and hold.cod_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("cod_seq").ToString) & "' " & _
                                        "and hold.coh_seq = '" & gU.dbEncode(coHolddt.Rows(j).Item("coh_seq").ToString) & "' " & _
                                        "and hold.COH_IN_STOCK_QTY <> hold.COH_QTY " & _
                                        "and hold.coh_type = 'STOCK' " & _
                                        "and hold.coh_status = 'HOLD' "
                            gDB.amendData(updateSql, gConn, transaction)
                        Next
                    Next

                    Dim ifPass As Boolean
                    Dim tempEDate As String = ""
                    Dim tempMDate As String = ""
                    Dim tempUOM2 As String = ""
                    Dim tempQty2 As String = ""

                    For Each rows As DataRow In pa_dt.Rows
                        If DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) > 0 OrElse DB.decodeDBNull(rows.Item("GRA_REJ_QTY"), 0) > 0 And rows.Item("mFlag").ToString <> "D" Then
                            For Each DROWS As DataRow In dt.Rows
                                If DROWS.Item("GRD_ITM_CODE").ToString.Trim = rows.Item("gra_itm_code").ToString.Trim AndAlso DROWS.Item("GRD_PACK_KEY").ToString.Trim = rows.Item("GRA_PACK_KEY").ToString.Trim AndAlso _
                                   DROWS.Item("GRD_PALLET_NO").ToString.Trim = rows.Item("GRA_PALLET_NO").ToString.Trim AndAlso DROWS.Item("GRD_BATCH_NO").ToString.Trim = rows.Item("GRA_BATCH_NO").ToString.Trim Then

                                    tempEDate = DROWS.Item("GRD_EXPIRY_DATE").ToString.Trim
                                    tempMDate = DROWS.Item("GRD_MANU_DATE").ToString.Trim

                                    tempUOM2 = DROWS.Item("GRD_UOM2").ToString.Trim
                                    tempQty2 = DROWS.Item("GRD_QTY2").ToString.Trim

                                    Exit For
                                End If
                            Next

                            st.IO_SEQ = ""
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNull(rows.Item("GRA_ITM_CODE").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("GRA_PACK_KEY").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_WH = gU.decodeNull(rows.Item("GRA_WH").ToString.Trim, "")
                            st.IO_AREA = ""
                            st.IO_LOC = gU.decodeNull(rows.Item("GRA_LOC").ToString.Trim, "")
                            st.IO_DOC = "GR"
                            st.IO_DOC_ID = GR_CODE.Text.Trim
                            st.IO_QTY = gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "")
                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, "")
                            st.IO_MANU_DATE = tempMDate
                            st.IO_EXPIRY_DATE = tempEDate

                            itmKey = gU.decodeNull(rows.Item("GRA_ITM_CODE").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("GRA_PACK_KEY").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "") & "#_#" & _
                                    gU.decodeNull(rows.Item("GRA_BATCH_NO").ToString.Trim, "")
                            If qtyDict.ContainsKey(itmKey) Then
                                Dim tmpCbm As Decimal = 0
                                Dim tmpwgt As Decimal = 0
                                Dim tmpqty As Decimal = 0

                                tmpCbm = cbmDict.Item(itmKey)
                                tmpwgt = wgtDict.Item(itmKey)
                                tmpqty = qtyDict.Item(itmKey)

                                If tmpqty = 0 Then tmpqty = 1

                                st.IO_CBM = Math.Round(tmpCbm / tmpqty * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                st.IO_KG = Math.Round(tmpwgt / tmpqty * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                'st.IO_CBM = Math.Round(tmpCbm * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                                'st.IO_KG = Math.Round(tmpwgt * CDbl(gU.decodeNull(rows.Item("GRA_PA_QTY").ToString.Trim, "0")), 2)
                            Else
                                st.IO_CBM = 0
                                st.IO_KG = 0
                            End If

                            st.IOS_DRUM_LEVEL = "1"
                            st.IOS_UOM2 = tempUOM2
                            st.IOS_ORG_QTY2 = tempQty2
                            st.IOS_QTY2 = tempQty2
                            st.IOS_SL = "N"
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                            'Loop to update serial
                            If snDt IsNot Nothing AndAlso snDt.Rows.Count > 0 Then

                                For k = 0 To snDt.Rows.Count - 1
                                    If snDt.Rows(k).Item("GRS_ITM_CODE").ToString.Trim = st.ITM_CODE AndAlso _
                                        snDt.Rows(k).Item("GRS_PACK_KEY").ToString.Trim = st.PACK_KEY AndAlso _
                                        gU.decodeNull(snDt.Rows(k).Item("GRS_PALLET_NO").ToString.Trim, "000") = st.PALLET_NO AndAlso _
                                        snDt.Rows(k).Item("GRS_BATCH_NO").ToString.Trim = st.lO_BATCH_NO AndAlso _
                                        snDt.Rows(k).Item("GRS_LOC").ToString.Trim = st.IO_LOC And snDt.Rows(k).Item("mFlag").ToString <> "D" Then

                                        st.IOS_DRUM_ID = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                        st.IOS_ORG_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim
                                        st.IOS_SERIAL_NO = snDt.Rows(k).Item("GRS_SERIAL_NO").ToString.Trim

                                        st.IO_QTY = 1
                                        ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)

                                        If Not ifPass Then
                                            Exit For
                                        End If

                                        ifPass = st.UnPostSERIAL(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)

                                        If Not ifPass Then
                                            Exit For
                                        End If

                                    End If
                                Next
                            Else
                                ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)

                                If Not ifPass Then
                                    Exit For
                                End If
                            End If


                            'Post for rejected item to damage zone
                            If CDbl(gU.decodeNullOrEmpty(rows.Item("GRA_REJ_QTY").ToString.Trim, "0")) > 0 Then
                                If dmgLoc = "" Then
                                    Throw New Exception("No damage location!")
                                Else
                                    st.IO_LOC = dmgLoc

                                    st.IO_QTY = gU.decodeNull(rows.Item("GRA_REJ_QTY").ToString.Trim, "")

                                    ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)

                                    If Not ifPass Then
                                        Exit For
                                    End If
                                End If
                            End If


                            selectSql = "select count(*) as value " & _
                                        "from WMS_REPLENISH_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                        "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " & _
                                        "AND ROD_BATCH_NO = '" & rows.Item("GRA_BATCH_NO").ToString.Trim & "' " & _
                                        "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' " & _
                                        "AND ROD_PALLET_NO = '" & gU.decodeNull(rows.Item("GRA_PALLET_NO").ToString.Trim, "000") & "' "

                            If DB.getValueFromSQL(selectSql, gConn, transaction) <> "0" Then
                                updateSql = "update WMS_REPLENISH_D " & _
                                            "set ROD_POST_QTY = ISNULL(ROD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                            "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " & _
                                            "AND ISNULL(ROD_BATCH_NO, '') = ISNULL('" & rows.Item("GRA_BATCH_NO").ToString.Trim & "', '') " & _
                                            "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' " & _
                                            "AND ROD_PALLET_NO = '" & rows.Item("GRA_PALLET_NO").ToString.Trim & "' "

                                gDB.amendData(updateSql, gConn, transaction)
                            Else
                                updateSql = "update WMS_REPLENISH_D " & _
                                            "set ROD_POST_QTY = ISNULL(ROD_POST_QTY,0) - " & DB.decodeDBNull(rows.Item("GRA_PA_QTY"), 0) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                            "AND ROD_ITM_CODE = '" & rows.Item("GRA_ITM_CODE").ToString.Trim & "' " & _
                                            "AND ROD_PACK_KEY = '" & rows.Item("GRA_PACK_KEY").ToString.Trim & "' "

                                gDB.amendData(updateSql, gConn, transaction)
                            End If
                        End If
                    Next

                    If ifPass Then

                        updateSql = "update WMS_REPLENISH " & _
                                    "set RO_STATUS = 'CLOSED', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                    "AND NOT EXISTS (" & _
                                        "SELECT 1 FROM WMS_REPLENISH_D D " & _
                                        "WHERE ISNULL(ROD_POST_QTY,0) <> ISNULL(ROD_QTY,0) " & _
                                        "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " & _
                                        "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " & _
                                        "AND D.RO_CODE = WMS_REPLENISH.RO_CODE) "

                        gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_REPLENISH " & _
                                    "set RO_STATUS = 'PAR', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and RO_CODE = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " & _
                                    "AND EXISTS (" & _
                                        "SELECT 1 FROM WMS_REPLENISH_D D " & _
                                        "WHERE ISNULL(ROD_POST_QTY,0) <> ISNULL(ROD_QTY,0) " & _
                                        "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " & _
                                        "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " & _
                                        "AND D.RO_CODE = WMS_REPLENISH.RO_CODE) "

                        'gDB.amendData(updateSql, gConn, transaction)

                        updateSql = "update WMS_GOODSRCV " & _
                                  "set GR_STATUS = 'NEW', " & _
                                      "GR_POSTED_DATE = Getdate(), " & _
                                      "GR_POSTED_BY = '" & Session("usr_id") & "', " & _
                                      "sys_lub = '" & Session("usr_id") & "', " & _
                                      "sys_lud = Getdate() " & _
                                  "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                  "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                  "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "

                        gDB.amendData(updateSql, gConn, transaction)


                        'Delete inspection records
                        deleteSql = "delete from WMS_GOODSRCV_INSP " & _
                                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text) & "' "

                        gDB.amendData(deleteSql, gConn, transaction)

                        transaction.Commit()

                        'No need BindGV as the page will be reload

                        'GR_STATUS.Text = "NEW"
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

                        'Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
                        'updtPnlLinkBar.Update()

                        'uiFun.displayMsgNew(updtPnlAlert, "1013", "", Session("gLang"))
                        successFlag = True
                    Else
                        transaction.Rollback()

                        successFlag = False

                        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
                        updtPnlLinkBar.Update()

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
                    errorMsg = ex.Message
                    'Throw ex
                End Try

                If successFlag = True Then
                    reloadPage("Record has been unPosted successfully!")

                    'Dim rmtPost As New RemotePost
                    'rmtPost.Url = "GRMain.aspx"
                    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    'rmtPost.Add("GR_CODE", GR_CODE.Text)
                    'rmtPost.alertMsg = "Record has been unPosted successfully!"
                    'rmtPost.Post()
                Else

                    If errorMsg <> "" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Un-Post Failed! " & errorMsg, Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
                    End If
                    'Dim rmtPost As New RemotePost
                    'rmtPost.Url = "GRMain.aspx"
                    'rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    'rmtPost.Add("GR_CODE", GR_CODE.Text)
                    'If errorMsg <> "" Then
                    '    rmtPost.alertMsg = "Un-Post Failed!\r\n" & errorMsg
                    'Else
                    '    rmtPost.alertMsg = "No item can be Un-Posted!"
                    'End If
                    'rmtPost.Post()
                End If

            End If
        Else
            Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
            updtPnlLinkBar.Update()

            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "沒有可供取消發布的物件!", Session("gLang"))
            End If
        End If
    End Sub

    Protected Sub cSBBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cSBBtn.Click
        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "GR", GR_CODE.Text, "../../")
        updtPnlLinkBar.Update()
    End Sub

    Private Sub SumCar(ByVal Carton As Double)
        TotalCar += Carton
    End Sub

    Private Sub SumKG(ByVal KG As Double, ByVal Qty As Integer)
        TotalKG += (KG * Qty)
    End Sub

    Private Sub SumCBM(ByVal CBM As Double)
        TotalCBM += CBM
    End Sub

    Private Function GetCarTotal() As Double
        Return TotalCar
    End Function

    Private Function GetKGTotal() As Double
        Return TotalKG
    End Function

    Private Function GetCBMTotal() As Double
        Return TotalCBM
    End Function

    Private Sub SumPCS(ByVal PCS As Double)
        SubTotalPCS += PCS
    End Sub

    Private Function GetTotal() As Double
        Return SubTotalPCS
    End Function

    Protected Sub btnItemDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnItemDelete.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim deleteArray As New ArrayList

            For x As Integer = 0 To dt.Rows.Count - 1
                If dt.Rows(x).Item("cb_select").ToString = "1" Then
                    Call ar.hideGVRow(GridView1, GridView1.Rows(x))

                    dt.Rows(x).Item("mFlag") = "D"

                    If dt.Rows(x).Item("grd_ref_seq").ToString = "" Then
                        deleteArray.Add(dt.Rows(x).Item("grd_seq").ToString)
                    End If

                    'dt.Rows(x).Item("cb_select") = "0"

                ElseIf dt.Rows(x).Item("grd_ref_seq").ToString <> "" Then
                    If deleteArray.Contains(dt.Rows(x).Item("grd_ref_seq").ToString) Then
                        dt.Rows(x).Item("mFlag") = "D"

                        Call ar.hideGVRow(GridView1, GridView1.Rows(x))
                    End If
                End If
            Next

            dt.AcceptChanges()

            Dim checkDt = From c In dt.Rows _
                          Where c.item("mFlag").ToString.Trim = "D" _
                           Select c.item("grd_itm_code")

            Dim checkList As List(Of Object) = checkDt.ToList

            If checkList.Count = dt.Rows.Count Then
                Dim result = (From c In dt.Rows _
                         Where c.item("mFlag") = "D" _
                          Select c.item("grd_itm_code")).Distinct

                Dim inList As List(Of Object) = result.ToList

                Dim pa_dt As DataTable
                pa_dt = Session("_M_IB_GR_TMP_pa_dt")


                For i As Integer = 0 To pa_dt.Rows.Count - 1
                    pa_dt.Rows(i).Item("mFlag") = "D"
                    pa_dt.Rows(i).AcceptChanges()
                Next
            End If
        End If
    End Sub

    Private Sub checkBarCode()
        Dim type As String = ""
        Dim ifMfgPN As Boolean = False
        Dim output As String = ""
        Dim popvalue As String = ""

        Dim tempCode As String = txtBarCode.Text

        Select Case STORER_CODE.SelectedValue
            Case "001"


                Select Case tempCode.Substring(0, 1).ToUpper.Trim
                    Case "D"
                        Dim tempDateStr As String = tempCode.Substring(1, tempCode.Length - 1).ToUpper.Trim
                        Dim tempDate As New Date

                        Try
                            tempDate = Date.ParseExact(tempDateStr, "yyyymmdd", CultureInfo.InvariantCulture)

                            type = "DATE"

                            output = Format(tempDate, "yyyy-MM-dd")

                        Catch ex As Exception
                            ifMfgPN = True
                        End Try

                    Case "L"
                        type = "L/N"
                        output = tempCode

                    Case "Q"
                        Dim tempQtyStr As String = tempCode.Substring(1, tempCode.Length - 1).ToUpper.Trim

                        If IsNumeric(tempQtyStr) Then
                            type = "QTY"
                            output = tempQtyStr
                        Else
                            ifMfgPN = True
                        End If

                    Case Else
                        ifMfgPN = True
                End Select

                If ifMfgPN Then
                    If tempCode.Length = 12 Then
                        If tempCode.Substring(4, 1).ToUpper.Trim = "B" Then
                            type = "P/N"
                            output = tempCode
                        End If
                    Else
                        Dim allNum As Boolean = True

                        For i As Integer = 0 To tempCode.Length - 1
                            If Not IsNumeric(tempCode.Chars(i)) Then
                                allNum = False
                                Exit For
                            End If
                        Next

                        If allNum Then
                            type = "PO"
                            output = tempCode
                        Else
                            If tempCode = "BUSSMANN" Then
                                type = "MFGNAME"
                                output = tempCode
                            Else
                                type = "MFGPN"
                                output = tempCode
                            End If
                        End If
                    End If

                End If

                Select Case type
                    Case "DATE"
                        popvalue = "Mfg D/C: " & output & ", " & "Orginal Value: " & txtBarCode.Text
                    Case "L/N"
                        popvalue = "L/N: " & output
                    Case "QTY"
                        popvalue = "QTY: " & output & ", " & "Orginal Value: " & txtBarCode.Text
                    Case "P/N"
                        popvalue = "P/N: " & output
                    Case "PO"
                        popvalue = "PO NO.: " & output
                    Case "MFGNAME"
                        popvalue = "Mfg Name: " & output
                    Case "MFGPN"
                        popvalue = "MFG P/N: " & output
                End Select
            Case "002"

                Dim result As String
                If tempCode.Contains("-") Then
                    result = "B/N"
                ElseIf tempCode.Length = 18 Then
                    result = "P/N"
                ElseIf Left(tempCode, 1) = "0" Then
                    result = "L/N or ID"
                ElseIf tempCode.Length = 8 Then
                    result = "D/C"
                ElseIf IsNumeric(tempCode) Then
                    result = "QTY"
                Else
                    If Not IsNumeric(Left(tempCode, 1)) And IsNumeric(Right(tempCode, tempCode.Length - 1)) Then
                        result = "Brand Code"
                    Else
                        result = "Brand Name"
                    End If
                End If

                Select Case result
                    Case "D/C"
                        Dim tempDate As New Date

                        Try
                            tempDate = Date.ParseExact(tempCode, "yyyymmdd", CultureInfo.InvariantCulture)

                            type = "DATE"

                            output = Format(tempDate, "yyyy-MM-dd")

                            popvalue = "Mfg D/C: " & output & ", " & "Orginal Value: " & tempCode
                        Catch ex As Exception
                            popvalue = "Mfg D/C: " & tempCode
                        End Try

                    Case "L/N or ID"
                        popvalue = "L/N: " & tempCode
                    Case "QTY"
                        popvalue = "QTY: " & tempCode
                    Case "P/N"
                        popvalue = "P/N: " & tempCode
                    Case "B/N"
                        popvalue = "B/N: " & tempCode
                    Case "QTY"
                        popvalue = "QTY: " & tempCode
                    Case "Brand Code"
                        popvalue = "Brand Code: " & tempCode
                    Case "Brand Name"
                        popvalue = "Brand Name: " & tempCode
                End Select
            Case Else
                popvalue = "Scanned Value: " & txtBarCode.Text
        End Select

        'ScriptManager.RegisterStartupScript(UpdatePanel1, GetType(UpdatePanel), "alertBarcode", "<script>alert('" & popvalue & "');</script>", True)

        If popvalue = "" Then popvalue = "Scanned Value: " & tempCode

        If popvalue <> "" Then
            'Me.ClientScript.RegisterStartupScript(Me.GetType, "alertBC", "<script>alert('" & popvalue & "');</script>")

            txtOutput.Text = popvalue
        End If

        txtBarCode.Text = ""
        txtBarCode.Focus()
    End Sub

    <System.Web.Services.WebMethod()> _
    Public Shared Function scanBarcode(ByVal val As String, ByVal storer As String) As String()
        Dim strArr(6) As String
        Dim type As String = ""
        Dim output As String = ""

        Dim tempCode As String = val
        Dim matched As Boolean = False

        strArr(0) = ""
        strArr(1) = ""
        strArr(2) = ""
        strArr(3) = ""
        strArr(4) = ""
        strArr(5) = ""
        strArr(6) = ""

        Select Case storer
            Case "001"
                Select Case tempCode.Substring(0, 1).ToUpper.Trim
                    Case "D"
                        Dim tempDateStr As String = tempCode.Substring(1, tempCode.Length - 1).ToUpper.Trim
                        Dim tempDate As New Date

                        Try
                            tempDate = Date.ParseExact(tempDateStr, "yyyyMMdd", CultureInfo.InvariantCulture)

                            type = "DATE"

                            output = Format(tempDate, "yyyy/MM/dd")

                        Catch ex As Exception
                            type = "P/N"
                            output = tempCode
                        End Try
                    Case "L"
                        type = "L/N"
                        output = tempCode

                    Case "Q"
                        Dim tempQtyStr As String = tempCode.Substring(1, tempCode.Length - 1).ToUpper.Trim

                        If IsNumeric(tempQtyStr) Then
                            type = "QTY"
                            output = tempQtyStr
                        Else
                            type = "P/N"
                            output = tempCode
                        End If

                    Case Else
                        Dim allNum As Boolean = True

                        For i As Integer = 0 To tempCode.Length - 1
                            If Not IsNumeric(tempCode.Chars(i)) Then
                                allNum = False
                                Exit For
                            End If
                        Next

                        If allNum Then
                            type = "PO"
                            output = tempCode
                        Else
                            If tempCode.Length = 12 Then
                                type = "P/N"
                                output = tempCode
                            Else
                                type = "MFGPN"
                                output = tempCode
                            End If
                        End If
                End Select

                Select Case type
                    Case "DATE"
                        strArr(3) = output
                        strArr(6) = "Date: " & tempCode
                    Case "QTY"
                        strArr(2) = output
                        strArr(6) = "Qty: " & tempCode
                    Case "P/N"
                        strArr(0) = output
                        strArr(6) = "P/N: " & tempCode
                    Case "PO"
                        strArr(1) = output
                        strArr(6) = "PO: " & tempCode
                    Case "L/N"
                        strArr(1) = output
                        strArr(6) = "L/N: " & tempCode
                    Case "MFGPN"
                        strArr(0) = output
                        strArr(6) = "Mfg P/N: " & tempCode
                    Case Else
                        strArr(6) = tempCode
                        strArr(6) = "Unrecognized: " & tempCode
                End Select

                If strArr(0) <> "" Then
                    Dim nDt As DataTable = HttpContext.Current.Session("dt")
                    For i As Integer = 0 To nDt.Rows.Count - 1
                        If nDt.Rows(i).Item("grd_itm_code").ToString.ToUpper = strArr(0).ToUpper Then
                            strArr(4) = i + 2
                            matched = True
                        End If
                    Next
                End If

            Case "002"
                Dim result As String

                If tempCode.Contains("-") And tempCode.Length = 10 Then
                    result = "B/N"
                ElseIf tempCode.Length = 8 Then
                    result = "D/C"
                ElseIf IsNumeric(tempCode) AndAlso tempCode.Substring(0, 1) <> "0" Then
                    result = "QTY"
                ElseIf Left(tempCode, 1) = "0" Then
                    result = "L/N"
                Else
                    result = "P/N"
                End If

                Select Case result
                    Case "D/C"
                        Dim tempDate As New Date

                        Try
                            tempDate = Date.ParseExact(tempCode, "yyyyMMdd", CultureInfo.InvariantCulture)

                            type = "DATE"

                            output = Format(tempDate, "yyyy/MM/dd")

                            strArr(3) = output

                        Catch ex As Exception
                            strArr(0) = output
                        End Try

                        strArr(6) = "Date: " & tempCode

                    Case "B/N"
                        strArr(1) = tempCode
                        strArr(6) = "B/N: " & tempCode
                    Case "QTY"
                        strArr(2) = tempCode
                        strArr(6) = "Qty: " & tempCode
                    Case "P/N"
                        strArr(0) = tempCode
                        strArr(6) = "P/N: " & tempCode
                    Case "L/N"
                        'strArr(5) = tempCode
                        strArr(6) = "L/N: " & tempCode
                    Case Else
                        strArr(6) = tempCode
                        strArr(6) = "Unrecognized: " & tempCode
                End Select

                If strArr(0) <> "" Then
                    Dim nDt As DataTable = HttpContext.Current.Session("dt")
                    For i As Integer = 0 To nDt.Rows.Count - 1
                        If nDt.Rows(i).Item("grd_itm_code").ToString.ToUpper = strArr(0).ToUpper Then
                            strArr(4) = i + 2
                            matched = True
                        End If
                    Next
                End If

            Case "003"
                Dim str2 As String()

                If tempCode.Contains(";") Then
                    str2 = tempCode.Trim.Split(";")

                    If str2.Length > 1 Then
                        If str2(0) <> "" Then
                            If str2(0).Length > 2 Then
                                If str2(0).Substring(0, 2) = "1P" Then
                                    Dim nDt As DataTable = HttpContext.Current.Session("dt")
                                    For i As Integer = 0 To nDt.Rows.Count - 1
                                        If nDt.Rows(i).Item("grd_itm_code").ToString.ToUpper = str2(0).Substring(2, str2(0).Length - 2).ToUpper Then
                                            strArr(4) = i + 2

                                            If str2(1).Length > 1 Then
                                                If str2(1).Substring(0, 1) = "Q" Then
                                                    strArr(2) = str2(1).Substring(1, str2(1).Length - 1)
                                                End If
                                            End If

                                            If str2(2).Length > 1 Then
                                                If str2(2).Substring(0, 2) = "1T" Then
                                                    strArr(1) = str2(2).Substring(2, str2(2).Length - 2)
                                                End If
                                            End If

                                            strArr(6) = "Item No./Qty/Lot No.: " & str2(0).Substring(2, str2(0).Length - 2) & "/" & strArr(2) & "/" & strArr(1)
                                            matched = True
                                        End If
                                    Next
                                Else
                                    strArr(0) = ""
                                    strArr(1) = ""
                                    strArr(2) = ""
                                    strArr(3) = ""
                                    strArr(4) = ""
                                    strArr(5) = ""
                                    strArr(6) = tempCode
                                End If
                            Else
                                strArr(0) = ""
                                strArr(1) = ""
                                strArr(2) = ""
                                strArr(3) = ""
                                strArr(4) = ""
                                strArr(5) = ""
                                strArr(6) = tempCode
                            End If
                        Else
                            strArr(0) = ""
                            strArr(1) = ""
                            strArr(2) = ""
                            strArr(3) = ""
                            strArr(4) = ""
                            strArr(5) = ""
                            strArr(6) = tempCode
                        End If
                    Else
                        strArr(0) = ""
                        strArr(1) = ""
                        strArr(2) = ""
                        strArr(3) = ""
                        strArr(4) = ""
                        strArr(5) = ""
                        strArr(6) = tempCode
                    End If
                ElseIf tempCode.Trim.Contains(" ") Then
                    str2 = tempCode.Trim.Split(" ")

                    If str2.Length > 1 Then
                        If str2(0) <> "" Then
                            Dim nDt As DataTable = HttpContext.Current.Session("dt")
                            For i As Integer = 0 To nDt.Rows.Count - 1
                                If nDt.Rows(i).Item("grd_itm_code").ToString.ToUpper = str2(0).ToUpper Then
                                    strArr(4) = i + 2
                                    strArr(2) = str2(1)

                                    strArr(6) = "Item/Qty: " & tempCode
                                    matched = True
                                End If
                            Next
                        End If

                        If matched = False Then
                            strArr(1) = str2(1)
                            strArr(5) = str2(0)

                            strArr(6) = "Batch/Lot: " & tempCode
                        End If
                    Else
                        strArr(0) = ""
                        strArr(1) = ""
                        strArr(2) = ""
                        strArr(3) = ""
                        strArr(4) = ""
                        strArr(5) = ""
                        strArr(6) = tempCode
                    End If
                Else
                    strArr(0) = ""
                    strArr(1) = ""
                    strArr(2) = ""
                    strArr(3) = ""
                    strArr(4) = ""
                    strArr(5) = ""
                    strArr(6) = tempCode
                End If


            Case Else
                strArr(0) = ""
                strArr(1) = ""
                strArr(2) = ""
                strArr(3) = ""
                strArr(4) = ""
                strArr(5) = ""
                strArr(6) = ""
        End Select

        Return strArr

    End Function

    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        ViewState("STORER_CODE") = STORER_CODE.SelectedValue
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

    Private Sub reloadPage(Optional ByVal alertMsg As String = "")
        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "RELOAD_PAGE", "reloadPage('" & gU.jsString(alertMsg) & "');", True)
    End Sub

    Protected Sub gvDetailSerialList_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvDetailSerialList.RowDataBound
        Dim paDt, tmpDt As DataTable
        Dim tmpRow As DataRow()

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                If DataBinder.Eval(e.Row.DataItem, "GRS_ITM_CODE").ToString.Trim = DS_ITEM_CODE.Text AndAlso _
                    DataBinder.Eval(e.Row.DataItem, "GRS_PACK_KEY").ToString.Trim = DS_PACK_KEY.Text AndAlso _
                    DataBinder.Eval(e.Row.DataItem, "GRS_BATCH_NO").ToString.Trim = DS_BATCH_NO.Text AndAlso _
                    DataBinder.Eval(e.Row.DataItem, "GRS_PALLET_NO").ToString.Trim = DS_PALLET_NO.Text AndAlso _
                    DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim <> "D" Then
                    e.Row.Visible = True
                Else
                    e.Row.Visible = False
                End If

                CType(e.Row.FindControl("GRS_SERIAL_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRS_SERIAL_NO").ToString.Trim

                'CType(e.Row.FindControl("DSP_GRS_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRS_LOC").ToString.Trim
                'CType(e.Row.FindControl("GRS_LOC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRS_LOC").ToString.Trim

                CType(e.Row.FindControl("GRS_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRS_SEQ").ToString.Trim

                paDt = Session("_M_IB_GR_TMP_pa_dt")

                tmpRow = paDt.Select("GRA_ITM_CODE = '" & gU.dbEncode(DS_ITEM_CODE.Text) & "' " & _
                                "AND GRA_PACK_KEY = '" & gU.dbEncode(DS_PACK_KEY.Text) & "' " & _
                                "AND isnull(GRA_BATCH_NO,'') = '" & gU.dbEncode(DS_BATCH_NO.Text) & "' " & _
                                "AND GRA_PALLET_NO = '" & gU.dbEncode(DS_PALLET_NO.Text) & "' " & _
                                "AND GRA_LOC IS NOT NULL", "GRA_LOC")

                If tmpRow.Count > 0 Then
                    tmpDt = paDt.Select("GRA_ITM_CODE = '" & gU.dbEncode(DS_ITEM_CODE.Text) & "' " & _
                                "AND GRA_PACK_KEY = '" & gU.dbEncode(DS_PACK_KEY.Text) & "' " & _
                                "AND isnull(GRA_BATCH_NO,'') = '" & gU.dbEncode(DS_BATCH_NO.Text) & "' " & _
                                "AND GRA_PALLET_NO = '" & gU.dbEncode(DS_PALLET_NO.Text) & "' " & _
                                "AND GRA_LOC IS NOT NULL", "GRA_LOC").CopyToDataTable

                    'tmpDt.Rows.Add()
                    'tmpDt.Rows(tmpDt.Rows.Count - 1).Item("GRA_LOC") = dmgLocValue

                    uiFun.load_dropdown(CType(e.Row.FindControl("GRS_LOC"), DropDownList), tmpDt, "GRA_LOC", "GRA_LOC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRS_LOC").ToString.Trim)

                    If tmpDt.Rows.Count = 1 AndAlso DataBinder.Eval(e.Row.DataItem, "GRS_LOC").ToString.Trim = "" Then
                        CType(e.Row.FindControl("GRS_LOC"), DropDownList).SelectedValue = tmpDt.Rows(0).Item("GRA_LOC").ToString.Trim
                    End If
                Else
                    CType(e.Row.FindControl("GRS_LOC"), DropDownList).Items.Insert(0, New ListItem("SELECT", ""))
                End If

                'Dim nImage As ImageButton = CType(e.Row.FindControl("Image_Loc_LookUp"), ImageButton)

                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp(document.myform." & HttpUtility.HtmlEncode(DS_WH.ClientID) & ".value, '" & _
                '                      HttpUtility.HtmlEncode(CType(e.Row.FindControl("DSP_GRS_LOC"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("GRS_LOC"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(DS_WH.ClientID) & "');return false;")
        End Select
    End Sub

    Protected Sub gvDetailSerialList_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvDetailSerialList.RowDeleting
        'Dim dtlDt As DataTable

        sn_dt = Session("sn_dt")

        If cU.gfBuildDataTableforGridView(sn_dt, gvDetailSerialList, True) Then
            If sn_dt.Rows(e.RowIndex).Item("mFlag").ToString.Trim = "N" Then
                sn_dt.Rows.RemoveAt(e.RowIndex)
            Else
                'Call ar.hideGVRow(gvAssVend, gvAssVend.Rows(e.RowIndex))
                'gvAssVend.Rows(e.RowIndex).Visible = False
                sn_dt.Rows(e.RowIndex).Item("mFlag") = "D"
            End If

            sn_dt.AcceptChanges()

            Session("sn_dt") = sn_dt

            gvDetailSerialList.DataSource = sn_dt
            gvDetailSerialList.DataBind()

            updtPnl_DetailSerialList.Update()
        End If
    End Sub

    Protected Sub btnAddDetailSerial_Click(sender As Object, e As System.EventArgs) Handles btnAddDetailSerial.Click
        Dim rowNo As Integer

        sn_dt = Session("sn_dt")

        If cU.gfBuildDataTableforGridView(sn_dt, gvDetailSerialList, True) Then
            sn_dt.Rows.Add()

            rowNo = sn_dt.Rows.Count - 1

            sn_dt.Rows(rowNo).Item("IMP_CODE") = IMP_CODE.Value.Trim
            sn_dt.Rows(rowNo).Item("STORER_CODE") = STORER_CODE.SelectedValue
            sn_dt.Rows(rowNo).Item("GR_CODE") = GR_CODE.Text.Trim
            sn_dt.Rows(rowNo).Item("GRS_ITM_CODE") = DS_ITEM_CODE.Text
            sn_dt.Rows(rowNo).Item("GRS_PACK_KEY") = DS_PACK_KEY.Text
            sn_dt.Rows(rowNo).Item("GRS_BATCH_NO") = DS_BATCH_NO.Text
            sn_dt.Rows(rowNo).Item("GRS_PALLET_NO") = DS_PALLET_NO.Text
            sn_dt.Rows(rowNo).Item("GRS_SERIAL_NO") = ""
            sn_dt.Rows(rowNo).Item("GRS_LOC") = ""

            sn_dt.Rows(rowNo).Item("mFlag") = "N"
            sn_dt.AcceptChanges()

            Session("sn_dt") = sn_dt

            gvDetailSerialList.DataSource = sn_dt
            gvDetailSerialList.DataBind()

            updtPnl_DetailSerialList.Update()
        End If

    End Sub

    Protected Sub btnDetailSerialOK_Click(sender As Object, e As System.EventArgs) Handles btnDetailSerialOK.Click
        sn_dt = Session("sn_dt")

        If cU.gfBuildDataTableforGridView(sn_dt, gvDetailSerialList, True) Then
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "HIDE_DETAIL_SERIAL", "$find('detailSerial_behavior').hide();", True)
        End If
    End Sub

    Private Sub updatePAParentChild(ByRef gConn As SqlConnection, ByRef transaction As SqlTransaction)
        Dim updateSql As String

        updateSql = "update WMS_GOODSRCV_PA " & _
                    "set GRA_SPLIT_FR = GRA_SEQ " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' "

        '"and GRA_SPLIT_FR is not null "

        gDB.amendData(updateSql, gConn, transaction)

        updateSql = "update WMS_GOODSRCV_PA " & _
                    "set GRA_SPLIT_FR = (" & _
                        "select convert(varchar, max(convert(int, p.GRA_SEQ))) " & _
                        "from WMS_GOODSRCV_PA p " & _
                        "where p.IMP_CODE = WMS_GOODSRCV_PA.IMP_CODE " & _
                        "and p.STORER_CODE = WMS_GOODSRCV_PA.STORER_CODE " & _
                        "and p.GR_CODE = WMS_GOODSRCV_PA.GR_CODE " & _
                        "and p.GRA_ITM_CODE = WMS_GOODSRCV_PA.GRA_ITM_CODE " & _
                        "and p.GRA_PACK_KEY = WMS_GOODSRCV_PA.GRA_PACK_KEY " & _
                        "and isnull(p.GRA_PALLET_NO, '000') = isnull(WMS_GOODSRCV_PA.GRA_PALLET_NO, '000') " & _
                        "and isnull(p.GRA_BATCH_NO, '') = isnull(WMS_GOODSRCV_PA.GRA_BATCH_NO, '') " & _
                        "and convert(int, p.GRA_SEQ) < convert(int, WMS_GOODSRCV_PA.GRA_SEQ)) " & _
                    "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "and GR_CODE = '" & gU.dbEncode(GR_CODE.Text.Trim) & "' " & _
                    "and exists (" & _
                        "select 1 from WMS_GOODSRCV_PA p " & _
                        "where p.IMP_CODE = WMS_GOODSRCV_PA.IMP_CODE " & _
                        "and p.STORER_CODE = WMS_GOODSRCV_PA.STORER_CODE " & _
                        "and p.GR_CODE = WMS_GOODSRCV_PA.GR_CODE " & _
                        "and p.GRA_ITM_CODE = WMS_GOODSRCV_PA.GRA_ITM_CODE " & _
                        "and p.GRA_PACK_KEY = WMS_GOODSRCV_PA.GRA_PACK_KEY " & _
                        "and isnull(p.GRA_PALLET_NO, '000') = isnull(WMS_GOODSRCV_PA.GRA_PALLET_NO, '000') " & _
                        "and isnull(p.GRA_BATCH_NO, '') = isnull(WMS_GOODSRCV_PA.GRA_BATCH_NO, '') " & _
                        "and convert(int, p.GRA_SEQ) < convert(int, WMS_GOODSRCV_PA.GRA_SEQ)) "

        gDB.amendData(updateSql, gConn, transaction)

    End Sub


    'Private Function getItemLocQty(ByVal itemCode As String, ByVal packKey As String, ByVal itemQty As Long, ByVal warehouse As String, _
    '                               Optional ByRef cnn As SqlConnection = Nothing, Optional ByRef transaction As SqlTransaction = Nothing) As List(Of itmLocQty)
    '    Dim selectSql As String
    '    Dim tmpDt As DataTable
    '    Dim tmpLocList As List(Of itmLocQty)
    '    Dim tmpItemLocQty As itmLocQty
    '    Dim IW_PREF_LOC1, IW_PREF_LOC2, LOC1_UTIL_TYPE, LOC2_UTIL_TYPE, ITEM_CBM, ITEM_KG, ITEM_AREA As String
    '    Dim LOC1_CBM, LOC2_CBM, LOC1_AREA, LOC2_AREA, LOC1_CBM_BAL, LOC1_KG_BAL, LOC1_AREA_BAL, LOC2_CBM_BAL, LOC2_KG_BAL, LOC2_AREA_BAL As String
    '    Dim utlLimit As Double = 0.75
    '    Dim checkLoc2 As Boolean = False
    '    Dim useLoc1, useLoc2 As Boolean

    '    tmpLocList = New List(Of itmLocQty)

    '    selectSql = "select iw.IW_PREF_LOC1, iw.IW_PREF_LOC2, " & _
    '                    "b1.BN_UTILIZATION_TYPE as LOC1_UTIL_TYPE, " & _
    '                    "b2.BN_UTILIZATION_TYPE as LOC2_UTIL_TYPE, " & _
    '                    "v.carton_cbm as item_cbm, v.aitm_vol as item_kg, v.aitm_length * v.aitm_width as item_area, " & _
    '                    "b1.BN_LENGTH * b1.BN_WIDTH * b1.BN_DEPTH / 1000000 as LOC1_CBM, " & _
    '                    "b1.BN_LENGTH * b1.BN_WIDTH * b1.BN_DEPTH / 1000000 as LOC2_CBM, " & _
    '                    "b1.BN_LENGTH * b1.BN_WIDTH / 10000 as LOC1_AREA, " & _
    '                    "b1.BN_LENGTH * b1.BN_WIDTH / 10000 as LOC2_AREA, " & _
    '                    "bal1.total_cbm as loc1_CBM_BAL, bal1.total_kg as loc1_kg_bal, bal1.total_area as loc1_area_bal, " & _
    '                    "bal2.total_cbm as loc2_CBM_BAL, bal2.total_kg as loc2_kg_bal, bal2.total_area as loc2_area_bal " & _
    '                "from WMS_ITEM_WH iw " & _
    '                "inner join v_alt_vend_item v " & _
    '                "on iw.ITM_CODE = v.ITM_CODE " & _
    '                "and iw.PACK_KEY = v.PACK_KEY " & _
    '                "left outer join V_LOCATION b1 " & _
    '                "on iw.IW_PREF_LOC1 = b1.loc " & _
    '                "left outer join V_LOCATION b2 " & _
    '                "on iw.IW_PREF_LOC2 = b2.loc " & _
    '                "left outer join ( " & _
    '                    "select b.ILOC_LOC, sum(i.carton_cbm * b.ILOC_BAL_QTY) as total_cbm, sum(i.aitm_vol * b.ILOC_BAL_QTY) as total_kg, sum(i.aitm_length * i.aitm_width * b.ILOC_BAL_QTY) as total_area " & _
    '                    "from WMS_ITEM_LOC_BAL b, v_alt_vend_item i " & _
    '                    "where b.imp_code = i.imp_code " & _
    '                    "and b.STORER_CODE = i.STORER_CODE " & _
    '                    "and b.ITM_CODE = i.ITM_CODE " & _
    '                    "and b.pack_key = i.PACK_KEY " & _
    '                    "group by b.ILOC_LOC) bal1 " & _
    '                "on iw.IW_PREF_LOC1 = bal1.ILOC_LOC " & _
    '                "left outer join ( " & _
    '                    "select b.ILOC_LOC, sum(i.carton_cbm * b.ILOC_BAL_QTY) as total_cbm, sum(i.aitm_vol * b.ILOC_BAL_QTY) as total_kg, sum(i.aitm_length * i.aitm_width * b.ILOC_BAL_QTY) as total_area " & _
    '                    "from WMS_ITEM_LOC_BAL b, v_alt_vend_item i " & _
    '                    "where b.imp_code = i.imp_code " & _
    '                    "and b.STORER_CODE = i.STORER_CODE " & _
    '                    "and b.ITM_CODE = i.ITM_CODE " & _
    '                    "and b.pack_key = i.PACK_KEY " & _
    '                    "group by b.ILOC_LOC) bal2 " & _
    '                "on iw.IW_PREF_LOC2 = bal2.ILOC_LOC " & _
    '                "where iw.ITM_CODE = '" & gU.dbEncode(itemCode) & "' " & _
    '                "and iw.PACK_KEY = '" & gU.dbEncode(packKey) & "' " & _
    '                "and iw.WH_CODE = '" & gU.dbEncode(warehouse) & "' "

    '    tmpDt = gDB.getDataTable(selectSql, cnn, transaction)

    '    If tmpDt.Rows.Count > 0 Then
    '        IW_PREF_LOC1 = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        IW_PREF_LOC2 = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_UTIL_TYPE = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_UTIL_TYPE = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        ITEM_CBM = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        ITEM_KG = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        ITEM_AREA = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_CBM = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_CBM = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_AREA = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_AREA = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_CBM_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_KG_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC1_AREA_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_CBM_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_KG_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
    '        LOC2_AREA_BAL = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim

    '        useLoc1 = False
    '        useLoc2 = False

    '        If IW_PREF_LOC1 <> "" AndAlso IW_PREF_LOC2 = "" Then
    '            useLoc1 = True
    '        ElseIf IW_PREF_LOC2 <> "" AndAlso IW_PREF_LOC1 = "" Then
    '            useLoc2 = True

    '        ElseIf IW_PREF_LOC2 = "" AndAlso IW_PREF_LOC1 = "" Then


    '        Else

    '            'If LOC1_UTIL_TYPE is empty, no need check ultilization, use prefer location 1
    '            If LOC1_UTIL_TYPE = "" Then
    '                useLoc1 = True

    '            ElseIf LOC1_UTIL_TYPE = "CBM" Then

    '                'If location 1 CBM is empty or item CBM is empt, no need check ultilization, use prefer location 1
    '                If LOC1_CBM = "" OrElse LOC1_CBM = "0" OrElse ITEM_CBM = "" OrElse ITEM_CBM = "0" Then
    '                    useLoc1 = True

    '                    'If item CBM + balance CBM <= location 1 utiliation limit, use prefer location 1
    '                ElseIf (LOC1_CBM - CDbl(ITEM_CBM) * itemQty + gU.decodeEmptyCdbl(LOC1_CBM_BAL, 0)) / LOC1_CBM <= utlLimit Then
    '                    useLoc1 = True
    '                End If

    '            ElseIf LOC1_UTIL_TYPE = "AREA" Then

    '                'If location 1 AREA is empty or item AREA is empt, no need check ultilization, use prefer location 1
    '                If LOC1_AREA = "" OrElse LOC1_AREA = "0" OrElse ITEM_AREA = "" OrElse ITEM_AREA = "0" Then
    '                    useLoc1 = True

    '                    'If item AREA + balance AREA <= location 1 utiliation limit, use prefer location 1
    '                ElseIf (LOC1_AREA - CDbl(ITEM_AREA) * itemQty + gU.decodeEmptyCdbl(LOC1_AREA_BAL, 0)) / LOC1_AREA <= utlLimit Then
    '                    useLoc1 = True
    '                End If

    '            End If

    '            If Not useLoc1 Then

    '                If LOC2_UTIL_TYPE = "CBM" Then
    '                    If LOC2_CBM = "" OrElse LOC2_CBM = "0" Then
    '                        useLoc2 = True

    '                        'If item CBM + balance CBM <= location 2 utiliation limit, use prefer location 2
    '                    ElseIf (LOC2_CBM - CDbl(ITEM_CBM) * itemQty + gU.decodeEmptyCdbl(LOC2_CBM_BAL, 0)) / LOC2_CBM <= utlLimit Then
    '                        useLoc2 = True
    '                    End If

    '                ElseIf LOC1_UTIL_TYPE = "AREA" Then

    '                    'If location 1 AREA is empty or item AREA is empt, no need check ultilization, use prefer location 1
    '                    If LOC1_AREA = "" OrElse LOC1_AREA = "0" OrElse ITEM_AREA = "" OrElse ITEM_AREA = "0" Then
    '                        useLoc2 = True

    '                        'If item AREA + balance AREA <= location 1 utiliation limit, use prefer location 1
    '                    ElseIf (LOC1_AREA - CDbl(ITEM_AREA) * itemQty + gU.decodeEmptyCdbl(LOC1_AREA_BAL, 0)) / LOC1_AREA <= utlLimit Then
    '                        useLoc2 = True
    '                    End If
    '                End If

    '                'If both location utilization is full, use back location 1
    '                'This part can implement more complex logic in the future
    '                If Not useLoc2 Then
    '                    useLoc1 = True
    '                End If

    '            End If
    '        End If


    '        If useLoc1 Then
    '            tmpItemLocQty = New itmLocQty
    '            tmpItemLocQty.loc = IW_PREF_LOC1
    '            tmpItemLocQty.qty = itemQty

    '            tmpLocList.Add(tmpItemLocQty)
    '        ElseIf useLoc2 Then
    '            tmpItemLocQty = New itmLocQty
    '            tmpItemLocQty.loc = IW_PREF_LOC2
    '            tmpItemLocQty.qty = itemQty

    '            tmpLocList.Add(tmpItemLocQty)
    '        End If

    '        Return tmpLocList
    '    Else
    '        Return tmpLocList
    '    End If

    'End Function


    'Private Sub genPutAway(ByRef grDt As DataTable, ByRef paDt As DataTable, Optional ByVal regenFlag As Boolean = True)
    '    Dim rows_count As Integer
    '    Dim i, j As Integer
    '    Dim plSeq As Integer
    '    Dim itmDict, paDict As Dictionary(Of String, itmStruct)
    '    Dim itmKey As String
    '    Dim keys As Dictionary(Of String, itmStruct).KeyCollection
    '    Dim keyArray As String()
    '    Dim tmpStruct As itmStruct
    '    Dim paLocList As List(Of itmLocQty)
    '    Dim delPaList As List(Of String)

    '    If regenFlag Then
    '        'Delete all existing put away
    '        For i = paDt.Rows.Count - 1 To 0 Step -1
    '            If paDt.Rows(i).Item("mFlag") = "N" Then
    '                paDt.Rows(i).Delete()
    '            Else
    '                Session("_M_IB_GR_TMP_pa_del_list") = gU.appendToList(Session("_M_IB_GR_TMP_pa_del_list"), paDt.Rows(i).Item("gra_seq"))
    '                paDt.Rows(i).Delete()
    '            End If
    '        Next
    '        paDt.AcceptChanges()

    '        'Set pa seq to zero
    '        Session("_M_IB_GR_TMP_pa_seq") = "0"

    '        plSeq = 0
    '    Else
    '        plSeq = Session("_M_IB_GR_TMP_pa_seq")
    '    End If

    '    If grDt.Rows.Count > 0 Then
    '        itmDict = New Dictionary(Of String, itmStruct)

    '        'Get items information
    '        For i = 0 To grDt.Rows.Count - 1
    '            If DB.decodeDBNull(grDt.Rows(i).Item("GRD_RCV_QTY"), 0) > 0 And grDt.Rows(i).Item("mFlag").ToString <> "D" Then
    '                itmKey = grDt.Rows(i).Item("grd_itm_code").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("grd_pack_key").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("grd_pallet_no").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("GRD_BATCH_NO").ToString

    '                If Not itmDict.ContainsKey(itmKey) Then
    '                    tmpStruct = New itmStruct
    '                    tmpStruct.itmName = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_ITM_NAME").ToString, "")
    '                    tmpStruct.itmSeq = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_SEQ").ToString, "")
    '                    tmpStruct.qty = gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)
    '                    tmpStruct.batchNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_BATCH_NO").ToString, "")
    '                    tmpStruct.skuNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

    '                    tmpStruct.expDate = grDt.Rows(i).Item("GRD_EXPIRY_DATE").ToString.Trim
    '                    tmpStruct.manuDate = grDt.Rows(i).Item("GRD_MANU_DATE").ToString.Trim

    '                    tmpStruct.whCode = grDt.Rows(i).Item("GRD_WH").ToString.Trim

    '                    itmDict.Add(itmKey, tmpStruct)
    '                Else
    '                    tmpStruct = itmDict.Item(itmKey)
    '                    tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)

    '                    itmDict(itmKey) = tmpStruct
    '                End If
    '            End If
    '        Next

    '        paDict = New Dictionary(Of String, itmStruct)

    '        delPaList = New List(Of String)

    '        If Not regenFlag Then
    '            'Get current put away information
    '            For i = 0 To paDt.Rows.Count - 1

    '                itmKey = paDt.Rows(i).Item("GRA_ITM_CODE").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_PACK_KEY").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_PALLET_NO").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_BATCH_NO").ToString

    '                If Not paDict.ContainsKey(itmKey) Then
    '                    tmpStruct = New itmStruct
    '                    tmpStruct.itmName = gU.decodeNullOrEmpty(paDt.Rows(i).Item("DSP_ITM_NAME").ToString, "")
    '                    tmpStruct.itmSeq = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_SPLIT_FR").ToString, "")
    '                    tmpStruct.qty = gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_PA_QTY"), 0)
    '                    tmpStruct.batchNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_BATCH_NO").ToString, "")
    '                    tmpStruct.skuNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

    '                    tmpStruct.expDate = paDt.Rows(i).Item("GRA_EXPIRY_DATE").ToString.Trim
    '                    tmpStruct.manuDate = paDt.Rows(i).Item("GRA_MANU_DATE").ToString.Trim

    '                    tmpStruct.whCode = paDt.Rows(i).Item("GRA_WH").ToString.Trim

    '                    paDict.Add(itmKey, tmpStruct)
    '                Else
    '                    tmpStruct = paDict.Item(itmKey)
    '                    tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_PA_QTY"), 0)

    '                    paDict(itmKey) = tmpStruct
    '                End If
    '            Next

    '            keys = paDict.Keys

    '            For i = 0 To keys.Count - 1
    '                If itmDict.ContainsKey(keys(i)) Then

    '                    tmpStruct = itmDict.Item(keys(i))

    '                    If tmpStruct.qty > paDict.Item(keys(i)).qty Then
    '                        tmpStruct.qty = tmpStruct.qty - paDict.Item(keys(i)).qty
    '                    Else
    '                        'Assume recv qty will not descrease
    '                        itmDict.Remove(keys(i))
    '                    End If
    '                Else
    '                    delPaList.Add(keys(i))
    '                End If
    '            Next
    '        End If


    '        keys = itmDict.Keys

    '        If keys.Count > 0 Then

    '            'Insert put away records
    '            For i = 0 To keys.Count - 1

    '                keyArray = Split(keys(i), "#_#")

    '                If Not regenFlag AndAlso paDict.ContainsKey(keys(i)) Then
    '                    'If item already exists in put away, increase the put away qty
    '                    For j = paDt.Rows.Count - 1 To 0 Step -1
    '                        itmKey = paDt.Rows(j).Item("GRA_ITM_CODE").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_PACK_KEY").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_PALLET_NO").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_BATCH_NO").ToString

    '                        If keys(i) = itmKey Then
    '                            paDt.Rows(j).Item("GRA_PA_QTY") = paDt.Rows(j).Item("GRA_PA_QTY") + itmDict.Item(keys(i)).qty

    '                            Exit For
    '                        End If
    '                    Next
    '                Else
    '                    'If it is a new item or re-gen put away, find the preferred location and add row
    '                    paLocList = getItemLocQty(keyArray(0), keyArray(1), itmDict.Item(keys(i)).qty, itmDict.Item(keys(i)).whCode)

    '                    If paLocList.Count > 0 Then
    '                        For j = 0 To paLocList.Count - 1
    '                            paDt.Rows.Add()

    '                            plSeq = plSeq + 1

    '                            rows_count = paDt.Rows.Count

    '                            paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
    '                            paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
    '                            paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
    '                            paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
    '                            paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
    '                            paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
    '                            paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
    '                            paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

    '                            paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = paLocList(j).qty
    '                            paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
    '                            paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
    '                            paDt.Rows(rows_count - 1).Item("GRA_LOC") = paLocList(j).loc
    '                            paDt.Rows(rows_count - 1).Item("mFlag") = "N"

    '                            paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

    '                            paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
    '                            paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

    '                            paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
    '                            paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
    '                        Next

    '                    Else
    '                        paDt.Rows.Add()

    '                        plSeq = plSeq + 1

    '                        rows_count = paDt.Rows.Count

    '                        paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
    '                        paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
    '                        paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
    '                        paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
    '                        paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
    '                        paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
    '                        paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
    '                        paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

    '                        paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = itmDict.Item(keys(i)).qty
    '                        paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
    '                        paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
    '                        paDt.Rows(rows_count - 1).Item("GRA_LOC") = ""
    '                        paDt.Rows(rows_count - 1).Item("mFlag") = "N"

    '                        paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

    '                        paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
    '                        paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

    '                        paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
    '                        paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
    '                    End If

    '                End If

    '            Next

    '            Session("_M_IB_GR_TMP_pa_seq") = CStr(plSeq)

    '            paDt.AcceptChanges()
    '        End If


    '    End If

    'End Sub


End Class
