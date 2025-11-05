Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text


Partial Class OUTBOUND_SO_SOMain
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

    Private dt As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT2 = gU.getConfig("DDFORMAT2")

        REM ****************************
        REM Modify Access Right Here

        ar = New AccessRightUtils("OB_SO", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME  from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(CUST_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
                                        "storer_code = '" & Session("usr_pref_storer") & "' " & _
                                      "and imp_code = '" & imp_code & "' " & _
                                      "ORDER BY 2", "CUS_CODE", "CUS_CODE", , Session("gSelectLabel"))

            uiFun.load_dropdownBy_ColCode(SH_MODE, "WMS_SHIP_ORDER.SH_MODE", Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            'btnConfirm.Visible = False
            If STORER_CODE.selectedValue = "" Then
                STORER_CODE.selectedValue = Session("usr_pref_storer")
            End If
            If SH_DATE.Text = "" Then
                SH_DATE.Text = Now.Date.ToString(DDFORMAT2)
            End If
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Sea Shipment Order Maintenance"
            lbl_ImageHd.Text = "Shipment Items"
            lbl_SH_CODE.Text = "Ship Order Code:"
            lbl_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_SH_DATE.Text = "Invoice Date:"
            lbl_CUST_CODE.Text = "Customer No.:"
            lbl_CUST_NAME.Text = "Customer Name:"
            lbl_SH_VEND_NO.Text = "Vendor No.:"
            lbl_SH_VEND_NAME.Text = "Vendor Name:"
            lbl_SH_MODE.Text = "Ship Mode:"
            lbl_SH_LC_BENEF.Text = "L/C Beneficiary:"
            lbl_SH_ETD.Text = "ETD:"
            lbl_SH_ETA.Text = "ETA:"
            lbl_SH_BOOKING_REF.Text = "Shipping Co. Booking Ref.:"
            lbl_SH_PORT_LOAD.Text = "Port Of Loading:"
            lbl_SH_PORT_DISCH.Text = "Port Of Discharge:"
            lbl_SH_FINAL_DEST.Text = "Final Destination:"
            lbl_SH_LC_NO.Text = "L/C Number:"
            lbl_SH_BILL_DESC.Text = "Billing Description:"
            lbl_SH_BL.Text = "Bill Of Landing:"
            lbl_SH_SEA_VESSEL.Text = "Vessel Name:"
            lbl_SH_AIR_FLIGHT.Text = "Flight:"
            lbl_SH_AIR_AWB.Text = "Airway Bill:"
            lbl_SH_AIR_HAWB.Text = "House Airway Bill:"
            lbl_SH_AIR_MAWB.Text = "Master Airway Bill:"
            lbl_SH_TARRIFF_CODE.Text = "Tariff Code:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"

            btnShip.Text = "Shipped"

            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel shipment?"");"
            btnShip.OnClientClick = "return confirm(""Are you sure to change shipment status to shipped?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this shipment?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this shipment?"");"

            If Session("pagemode") = "N" Then
                SH_CODE.Text = "[No. will be auto generated]"
            End If
        ElseIf Session("gLang") = "C" Then
            lheader.Value = "裝運指令資料庫"
            lbl_ImageHd.Text = "裝運物品"
            lbl_SH_CODE.Text = "裝運指令編碼:"
            lbl_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_SH_DATE.Text = "發票日期:"
            lbl_CUST_CODE.Text = "客戶編號:"
            lbl_CUST_NAME.Text = "客戶名稱:"
            lbl_SH_VEND_NO.Text = "供應商編號:"
            lbl_SH_VEND_NAME.Text = "供應商名稱:"

            lbl_SH_MODE.Text = "裝運模式:"
            lbl_SH_LC_BENEF.Text = "L/C受益人:"
            lbl_SH_ETD.Text = "預計交貨時間:"
            lbl_SH_ETA.Text = "預計到達時間:"
            lbl_SH_BOOKING_REF.Text = "船務公司預訂號.:"
            lbl_SH_PORT_LOAD.Text = "裝貨港:"
            lbl_SH_PORT_DISCH.Text = "卸貨港:"

            lbl_SH_LC_NO.Text = "L/C編號:"
            lbl_SH_BILL_DESC.Text = "提單說明:"
            lbl_SH_BL.Text = "海運提單:"
            lbl_SH_SEA_VESSEL.Text = "船名:"
            lbl_SH_AIR_FLIGHT.Text = "航班:"
            lbl_SH_AIR_AWB.Text = "空運單:"
            lbl_SH_AIR_HAWB.Text = "空運提單:"
            lbl_SH_AIR_MAWB.Text = "主空運單:"
            lbl_SH_TARRIFF_CODE.Text = "關稅代碼:"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"

            btnShip.Text = "運送"

            CancelBtn.OnClientClick = "return confirm(""確定取消裝運指令?"");"
            btnShip.OnClientClick = "return confirm(""確定運送裝運指令?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存指令?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存指令?"");"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"

        End If
        REM **********************

        REM **********************
        REM Additional CSS
        SH_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            STORER_CODE.CssClass = "REQUIRED"
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("SH_CODE") = ""

            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        If moduleAction = "SELECTDO" Then
            Call addItemtoDO()
        End If

        Call changeLabel()

        Dim colIdx_StartWith As Integer = 0

        ar.addColDef("sd_disp_seq", "sd_disp_seq", colIdx_StartWith)
        ar.addColDef("sd_pallet_no", "sd_pallet_no", colIdx_StartWith)
        ar.addColDef("sd_carton_no", "sd_carton_no", colIdx_StartWith)
        ar.addColDef("sd_itm_code", "sd_itm_code", colIdx_StartWith)
        ar.addColDef("sd_pack_key", "sd_pack_key", colIdx_StartWith)
        ar.addColDef("sd_itm_desc", "sd_itm_desc", colIdx_StartWith)
        ar.addColDef("sd_pack_no", "sd_pack_no", colIdx_StartWith)
        ar.addColDef("sd_pack_type", "sd_pack_type", colIdx_StartWith)
        ar.addColDef("sd_qty", "sd_qty", colIdx_StartWith)
        ar.addColDef("sd_uom", "sd_uom", colIdx_StartWith)
        ar.addColDef("sd_pcs_uom", "sd_pcs_uom", colIdx_StartWith)
        ar.addColDef("sd_totpcs", "sd_totpcs", colIdx_StartWith)
        ar.addColDef("sd_tot_wgt", "sd_tot_wgt", colIdx_StartWith)
        ar.addColDef("sd_tot_cbm", "sd_tot_cbm", colIdx_StartWith)
        ar.addColDef("sd_rem", "sd_rem", colIdx_StartWith)

        cm = New CommonMenu("SO", lheader.Value, SH_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "Y"
        cm.haveAttachments = "Y"
        cm.haveNotes = "Y"
        cm.haveTasks = "Y"
        cm.haveEmail = "Y"
        cm.haveHistory = "Y"

        cm.genCM(cmBar)

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "SO", SH_CODE.Text, "../../")

        REM Select DO button
        selectItemBtn.Attributes.Add("onclick", "DOLookUp(document.myform." & STORER_CODE.ClientID & ".value);")

        If STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            If STATUS.Text = "SHIPPED" Then btnShip.Visible = False Else btnShip.Enabled = False
        ElseIf STATUS.Text = "SHIPPED" Then
            ar.sec_viewMode = "Y"
            'ar.sec_write = "N"
            btnShip.Visible = False
        End If

        btnVnd_LookUp.Attributes.Add("onclick", "javascript:VendorLookUp()")

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
        ar.hideGVForStorer(GridView1, STORER_CODE.Text, "SO", "WMS_SHIP_ORDER_D")
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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("sd_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_disp_seq").ToString.Trim
                CType(e.Row.FindControl("sd_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_pallet_no").ToString.Trim
                CType(e.Row.FindControl("sd_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_carton_no").ToString.Trim
                CType(e.Row.FindControl("sd_itm_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_itm_code").ToString.Trim
                CType(e.Row.FindControl("sd_pack_key"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_pack_key").ToString.Trim
                CType(e.Row.FindControl("sd_itm_desc"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_itm_desc").ToString.Trim
                CType(e.Row.FindControl("sd_pack_key"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_pack_no").ToString.Trim
                CType(e.Row.FindControl("sd_pack_type"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_pack_type").ToString.Trim
                CType(e.Row.FindControl("sd_track_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_track_no").ToString.Trim
                CType(e.Row.FindControl("sd_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "sd_qty").ToString.Trim)

                Dim nDropDown As DropDownList = CType(e.Row.FindControl("sd_uom"), DropDownList)
                uiFun.load_dropdown(nDropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(imp_code) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "sd_uom").ToString.Trim
                CType(e.Row.FindControl("sd_uom"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "sd_uom").ToString.Trim

                CType(e.Row.FindControl("sd_pcs_uom"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "sd_pcs_uom").ToString.Trim)
                CType(e.Row.FindControl("sd_totpcs"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "sd_totpcs").ToString.Trim)
                CType(e.Row.FindControl("sd_tot_wgt"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "sd_tot_wgt").ToString.Trim)
                CType(e.Row.FindControl("sd_tot_cbm"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "sd_tot_cbm").ToString.Trim)
                CType(e.Row.FindControl("sd_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "sd_rem").ToString.Trim
                REM **********************

                CType(e.Row.FindControl("sd_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("sd_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("sd_pcs_uom"), TextBox).ClientID & ".value * this.value")
                CType(e.Row.FindControl("sd_pcs_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("sd_totpcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("sd_qty"), TextBox).ClientID & ".value * this.value")

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                'nButton.ID = "btnDelete" & e.Row.RowIndex

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

                ar.hideGVForStorer(GridView1, STORER_CODE.Text, "SO", "WMS_SHIP_ORDER_D", e)
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select max(Convert(INT,SD_SEQ)) + 1 from wms_ship_order_d " & _
                                        "where imp_code = '" & gU.dbEncode(imp_code.Trim.ToString) & "' " & _
                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and sh_code = '" & gU.dbEncode(SH_CODE.Text) & "' "
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
            dt.Rows(rows_count - 1).Item("sd_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        'If SH_CODE.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_SH_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_SH_CODE.Text & "不能空白!", Session("gLang"))
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

        If SH_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_SH_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_SH_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(SH_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_SH_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_SH_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If


        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1

                If uiFun.gvValidate(Me, dt, "sd_qty", "Qty", _
                                 CType(GridView1.Rows(i).FindControl("sd_qty"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "sd_pcs_uom", "PCS/ UOM", _
                                 CType(GridView1.Rows(i).FindControl("sd_pcs_uom"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "sd_totpcs", "Total PCS", _
                                 CType(GridView1.Rows(i).FindControl("sd_totpcs"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "sd_tot_wgt", "Total Wgt", _
                                 CType(GridView1.Rows(i).FindControl("sd_tot_wgt"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "sd_tot_cbm", "Total CBM", _
                                 CType(GridView1.Rows(i).FindControl("sd_tot_cbm"), TextBox).Text) = False Then Return False

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

        Dim dupSQL As String
        Dim dupTbl As New DataTable

        If validateAll() Then

            Select Case SH_MODE.SelectedValue
                Case "SEA"
                    SH_AIR_FLIGHT.Text = ""
                    SH_AIR_AWB.Text = ""
                    SH_AIR_MAWB.Text = ""
                    SH_AIR_HAWB.Text = ""
                Case "AIR"
                    SH_BL.Text = ""
                    SH_SEA_VESSEL.Text = ""
                Case Else
                    SH_BL.Text = ""
                    SH_SEA_VESSEL.Text = ""
                    SH_AIR_FLIGHT.Text = ""
                    SH_AIR_AWB.Text = ""
                    SH_AIR_MAWB.Text = ""
                    SH_AIR_HAWB.Text = ""
            End Select

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("SO", gConn, transaction)
                    'nextNo = SH_CODE.Text
                    REM **********************

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    dupSQL = "select 1 from wms_ship_order " & _
                            "where sh_code = '" & gU.dbEncode(nextNo) & "' " & _
                            "and imp_code='" & gU.dbEncode(imp_code) & "' " & _
                            "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_ship_order ( " & _
                                        "sh_code,imp_code,storer_code, " & _
                                        "status,cust_code, " & _
                                        "sh_vend_no,sh_vend_name,sh_date, " & _
                                        "sh_lc_no,sh_lc_benef,sh_bill_desc, " & _
                                        "sh_mode,sh_sea_vessel,sh_air_flight, " & _
                                        "sh_etd,sh_eta,sh_bl, " & _
                                        "sh_air_awb,sh_air_hawb,sh_air_mawb, " & _
                                        "sh_port_load,sh_port_disch,sh_final_dest, " & _
                                        "sh_booking_ref,sh_tarriff_code, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values ( " & _
                                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(STATUS.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUST_CODE.SelectedValue)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(H_SH_VEND_NO.Value)) & "," & gU.convdbNVCData(gU.dbEncode(H_SH_VEND_NAME.Value)) & "," & gU.convdbDate(gU.dbEncode(SH_DATE.Text)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(SH_LC_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_LC_BENEF.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_BILL_DESC.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(SH_MODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(SH_SEA_VESSEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_AIR_FLIGHT.Text)) & ", " & _
                                        gU.convdbDate(gU.dbEncode(SH_ETD.Text)) & "," & gU.convdbDate(gU.dbEncode(SH_ETA.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_BL.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(SH_AIR_AWB.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_AIR_HAWB.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_AIR_MAWB.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(SH_PORT_LOAD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_PORT_DISCH.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_FINAL_DEST.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(SH_BOOKING_REF.Text)) & "," & gU.convdbNVCData(gU.dbEncode(SH_TARRIFF_CODE.Text)) & "," & _
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "sd_disp_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into wms_ship_order_d ( " & _
                                                   "sh_code, imp_code, storer_code, sd_seq, sd_disp_seq, " & _
                                                   "sd_pallet_no, sd_carton_no, sd_itm_code, " & _
                                                   "sd_pack_key, sd_itm_desc, sd_pack_no, " & _
                                                   "sd_pack_type, sd_qty, sd_uom, sd_pcs_uom, " & _
                                                   "sd_totpcs, sd_tot_wgt, sd_tot_cbm, sd_rem, sd_track_no, do_code, " & _
                                                   "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                   "values ( " & _
                                                   gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_seq").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_disp_seq").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pallet_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_carton_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_code").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_key").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_desc").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_type").ToString.Trim, ""))) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_qty").ToString.Trim, "0")) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_uom").ToString.Trim, ""))) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_pcs_uom").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_totpcs").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_wgt").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_cbm").ToString.Trim, "0")) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_rem").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_track_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("do_code").ToString.Trim, ""))) & ", " & _
                                                   "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in " & lheader.Value & "!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", lheader.Value & "資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_ship_order set " & _
                                    "status = " & gU.convdbNVCData(gU.dbEncode(STATUS.Text)) & ", " & _
                                    "cust_code = " & gU.convdbNVCData(gU.dbEncode(CUST_CODE.SelectedValue)) & ", " & _
                                    "sh_vend_no = " & gU.convdbNVCData(gU.dbEncode(H_SH_VEND_NO.Value)) & ", " & _
                                    "sh_vend_name = " & gU.convdbNVCData(gU.dbEncode(H_SH_VEND_NAME.Value)) & ", " & _
                                    "sh_date = " & gU.convdbDate(gU.dbEncode(SH_DATE.Text)) & ", " & _
                                    "sh_lc_no = " & gU.convdbNVCData(gU.dbEncode(SH_LC_NO.Text)) & ", " & _
                                    "sh_lc_benef = " & gU.convdbNVCData(gU.dbEncode(SH_LC_BENEF.Text)) & ", " & _
                                    "sh_bill_desc = " & gU.convdbNVCData(gU.dbEncode(SH_BILL_DESC.Text)) & ", " & _
                                    "sh_mode = " & gU.convdbNVCData(gU.dbEncode(SH_MODE.SelectedValue)) & ", " & _
                                    "sh_sea_vessel = " & gU.convdbNVCData(gU.dbEncode(SH_SEA_VESSEL.Text)) & ", " & _
                                    "sh_air_flight = " & gU.convdbNVCData(gU.dbEncode(SH_AIR_FLIGHT.Text)) & ", " & _
                                    "sh_etd = " & gU.convdbDate(gU.dbEncode(SH_ETD.Text)) & ", " & _
                                    "sh_eta = " & gU.convdbDate(gU.dbEncode(SH_ETA.Text)) & ", " & _
                                    "sh_bl = " & gU.convdbNVCData(gU.dbEncode(SH_BL.Text)) & ", " & _
                                    "sh_air_awb = " & gU.convdbNVCData(gU.dbEncode(SH_AIR_AWB.Text)) & ", " & _
                                    "sh_air_hawb = " & gU.convdbNVCData(gU.dbEncode(SH_AIR_HAWB.Text)) & ", " & _
                                    "sh_air_mawb = " & gU.convdbNVCData(gU.dbEncode(SH_AIR_MAWB.Text)) & ", " & _
                                    "sh_port_load = " & gU.convdbNVCData(gU.dbEncode(SH_PORT_LOAD.Text)) & ", " & _
                                    "sh_port_disch = " & gU.convdbNVCData(gU.dbEncode(SH_PORT_DISCH.Text)) & ", " & _
                                    "sh_final_dest = " & gU.convdbNVCData(gU.dbEncode(SH_FINAL_DEST.Text)) & ", " & _
                                    "sh_booking_ref = " & gU.convdbNVCData(gU.dbEncode(SH_BOOKING_REF.Text)) & ", " & _
                                    "sh_tarriff_code = " & gU.convdbNVCData(gU.dbEncode(SH_TARRIFF_CODE.Text)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and SH_code = '" & gU.dbEncode(SH_CODE.Text) & "' "

                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "sd_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_ship_order_d ( " & _
                                                   "sh_code, imp_code, storer_code, sd_seq, sd_disp_seq, " & _
                                                   "sd_pallet_no, sd_carton_no, sd_itm_code, " & _
                                                   "sd_pack_key, sd_itm_desc, sd_pack_no, " & _
                                                   "sd_pack_type, sd_qty, sd_uom, sd_pcs_uom, " & _
                                                   "sd_totpcs, sd_tot_wgt, sd_tot_cbm, sd_rem, sd_track_no, do_code, " & _
                                                   "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                   "values ( " & _
                                                   gU.convdbNVCData(gU.dbEncode(SH_CODE.Text.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_seq").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_disp_seq").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pallet_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_carton_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_code").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_key").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_desc").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_type").ToString.Trim, ""))) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_qty").ToString.Trim, "0")) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_uom").ToString.Trim, ""))) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_pcs_uom").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_totpcs").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_wgt").ToString.Trim, "0")) & ", " & _
                                                   gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_cbm").ToString.Trim, "0")) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_rem").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_track_no").ToString.Trim, ""))) & ", " & _
                                                   gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("do_code").ToString.Trim, ""))) & ", " & _
                                                   "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_ship_order_d " & _
                                            "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and sh_code = '" & gU.dbEncode(SH_CODE.Text.Trim) & "' " & _
                                            "and sd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("sd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update wms_ship_order_d set " & _
                                                "sd_disp_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_disp_seq").ToString.Trim, ""))) & ", " & _
                                                "sd_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pallet_no").ToString.Trim, ""))) & ", " & _
                                                "sd_carton_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_carton_no").ToString.Trim, ""))) & ", " & _
                                                "sd_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_code").ToString.Trim, ""))) & ", " & _
                                                "sd_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_key").ToString.Trim, ""))) & ", " & _
                                                "sd_itm_desc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_itm_desc").ToString.Trim, ""))) & ", " & _
                                                "sd_pack_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_no").ToString.Trim, ""))) & ", " & _
                                                "sd_pack_type = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_pack_type").ToString.Trim, ""))) & ", " & _
                                                "sd_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_qty").ToString.Trim, "0")) & ", " & _
                                                "sd_uom = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_uom").ToString.Trim, ""))) & ", " & _
                                                "sd_pcs_uom = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_pcs_uom").ToString.Trim, "0")) & ", " & _
                                                "sd_totpcs = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_totpcs").ToString.Trim, "0")) & ", " & _
                                                "sd_tot_wgt = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_wgt").ToString.Trim, "0")) & ", " & _
                                                "sd_tot_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("sd_tot_cbm").ToString.Trim, "0")) & ", " & _
                                                "sd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_rem").ToString.Trim, ""))) & ", " & _
                                                "sd_track_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("sd_track_no").ToString.Trim, ""))) & ", " & _
                                                "do_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("do_code").ToString.Trim, ""))) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                                "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                "and sh_code = '" & gU.dbEncode(SH_CODE.Text.Trim) & "' " & _
                                                "and sd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("sd_seq").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    SH_CODE.Text = nextNo
                    ViewState("SH_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    SH_CODE.ForeColor = Drawing.Color.Black
                    SH_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))

                Call BindGV()

            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If
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

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim pk_code, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("SH_CODE") <> "" Then
            pk_code = ViewState("SH_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("SH_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            SH_CODE.ForeColor = Drawing.Color.Red
            STATUS.Text = "NEW"
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT wms_ship_order.IMP_CODE,  wms_ship_order.STORER_CODE,  wms_ship_order.SH_CODE, " & _
                        "  wms_ship_order.STATUS,  wms_ship_order.CUST_CODE,  wms_ship_order.SH_VEND_NO, " & _
                        "  wms_ship_order.SH_VEND_NAME, CONVERT(varchar, (wms_ship_order.SH_DATE," & DDFORMAT & ") as SH_DATE,  wms_ship_order.SH_LC_NO, " & _
                        "  wms_ship_order.SH_LC_BENEF,  wms_ship_order.SH_BILL_DESC,  wms_ship_order.SH_MODE, " & _
                        "  wms_ship_order.SH_SEA_VESSEL,  wms_ship_order.SH_AIR_FLIGHT,  CONVERT(varchar, (wms_ship_order.SH_ETD," & DDFORMAT & ") as SH_ETD, " & _
                        "  CONVERT(varchar, (wms_ship_order.SH_ETA," & DDFORMAT & ") as SH_ETA ,  wms_ship_order.SH_BL,  wms_ship_order.SH_AIR_AWB, " & _
                        "  wms_ship_order.SH_AIR_HAWB,  wms_ship_order.SH_AIR_MAWB,  wms_ship_order.SH_PORT_LOAD, " & _
                        "  wms_ship_order.SH_PORT_DISCH,  wms_ship_order.SH_FINAL_DEST,  wms_ship_order.SH_BOOKING_REF, " & _
                        "  wms_ship_order.SH_TARRIFF_CODE,  wms_ship_order.SYS_LUB,  wms_ship_order.SYS_LUD, " & _
                        "  wms_ship_order.SYS_CD,  wms_ship_order.SYS_CB " & _
                        " FROM wms_ship_order " & _
                        "where wms_ship_order.sh_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and wms_ship_order.imp_code = '" & imp_code & "' " & _
                        "and wms_ship_order.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                'imp_code = dt.Rows(0).Item("IMP_CODE").ToString

                SH_CODE.Text = dt.Rows(0).Item("SH_CODE").ToString
                STATUS.Text = dt.Rows(0).Item("STATUS").ToString

                SH_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("SH_DATE").ToString)
                SH_VEND_NO.Text = dt.Rows(0).Item("SH_VEND_NO").ToString
                SH_VEND_NAME.Text = dt.Rows(0).Item("SH_VEND_NAME").ToString
                H_SH_VEND_NO.Value = dt.Rows(0).Item("SH_VEND_NO").ToString
                H_SH_VEND_NAME.Value = dt.Rows(0).Item("SH_VEND_NAME").ToString
                SH_MODE.SelectedValue = dt.Rows(0).Item("SH_MODE").ToString
                SH_LC_BENEF.Text = dt.Rows(0).Item("SH_LC_BENEF").ToString
                SH_ETD.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("SH_ETD").ToString)
                SH_ETA.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("SH_ETA").ToString)
                SH_BOOKING_REF.Text = dt.Rows(0).Item("SH_BOOKING_REF").ToString
                SH_PORT_LOAD.Text = dt.Rows(0).Item("SH_PORT_LOAD").ToString
                SH_PORT_DISCH.Text = dt.Rows(0).Item("SH_PORT_DISCH").ToString
                SH_LC_NO.Text = dt.Rows(0).Item("SH_LC_NO").ToString
                SH_BILL_DESC.Text = dt.Rows(0).Item("SH_BILL_DESC").ToString
                SH_BL.Text = dt.Rows(0).Item("SH_BL").ToString
                SH_SEA_VESSEL.Text = dt.Rows(0).Item("SH_SEA_VESSEL").ToString
                SH_AIR_FLIGHT.Text = dt.Rows(0).Item("SH_AIR_FLIGHT").ToString
                SH_AIR_AWB.Text = dt.Rows(0).Item("SH_AIR_AWB").ToString
                SH_AIR_HAWB.Text = dt.Rows(0).Item("SH_AIR_HAWB").ToString
                SH_AIR_MAWB.Text = dt.Rows(0).Item("SH_AIR_MAWB").ToString
                SH_TARRIFF_CODE.Text = dt.Rows(0).Item("SH_TARRIFF_CODE").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                CUST_NAME.Text = DB.getValueFromSQL("SELECT CUS_NAME FROM WMS_CUSTOMER WHERE CUS_CODE = '" & gU.dbEncode(dt.Rows(0).Item("CUST_CODE").ToString) & "' ")

                uiFun.load_dropdown(CUST_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
                                              "storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                              "and imp_code = '" & imp_code & "' " & _
                                              "ORDER BY 2", "CUS_CODE", "CUS_CODE", , Session("gSelectLabel"))

                CUST_CODE.SelectedValue = dt.Rows(0).Item("CUST_CODE").ToString
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        SQLString = "SELECT wms_ship_order_d.*, 'U' as mFlag from wms_ship_order_d " & _
                    "where wms_ship_order_d.SH_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_ship_order_d.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_ship_order_d.imp_code = '" & imp_code & "'"

        SQLString = SQLString & " order by wms_ship_order_d.sd_disp_seq"
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
        STATUS.Text = "CANCELLED"
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

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Protected Sub CUST_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUST_CODE.SelectedIndexChanged
        Dim SQLSTR As String = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " & _
                                "cus_area_del, cus_region_del, cus_country_del, cus_cont_per_ord, cus_cont_tel_ord " & _
                                "from wms_customer where cus_code = '" & CUST_CODE.SelectedValue & "' " & _
                                "and storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                                "and imp_code = '" & imp_code & "'"
        Dim cust_table As New DataTable

        cust_table = gDB.getDataTable(SQLSTR)

        If cust_table.Rows.Count > 0 Then
            CUST_NAME.Text = cust_table.Rows(0).Item("cus_name").ToString

        Else
            CUST_NAME.Text = ""
        End If
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        CUST_NAME.Text = ""

        Session("gSelectLabel") = "SELECT"
        uiFun.load_dropdown(CUST_CODE, "select CUS_CODE, CUS_NAME from wms_customer where " & _
                   " storer_code = '" & STORER_CODE.SelectedValue & "' " & _
                 "and imp_code = '" & imp_code & "' " & _
                 "ORDER BY 2", "CUS_CODE", "CUS_CODE", , Session("gSelectLabel"))


    End Sub

    Protected Sub addItemtoDO()
        Dim i As Integer

        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(Convert(int,SD_SEQ)) + 1 from WMS_SHIP_ORDER_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(imp_code) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and SH_CODE = '" & gU.dbEncode(SH_CODE.Text.Trim) & "' "
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
            Dim do_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim seqListarray As String()

            itemListarray = Split(itemList.Value, ", ")
            seqListarray = Split(seqList.Value, ", ")
            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = itemList.Value & "_000_" & seqList.Value
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, itemListarray(i) & "_000_" & seqListarray(i))
                Next
            End If

            SQLString = "SELECT d.*, m.* " & _
            "from WMS_DELV_ORDER_D d, WMS_DELV_ORDER m " & _
            "where m.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            "and m.IMP_CODE = '" & gU.dbEncode(imp_code) & "' " & _
            "AND M.DO_CODE + '_000_' + D.DOD_SEQ IN ('" & Replace(itemPackList, ", ", "', '") & "') " & _
            "AND M.DO_CODE = D.DO_CODE " & _
            "AND M.IMP_CODE = D.IMP_CODE " & _
            "AND M.STORER_CODE = D.STORER_CODE "

            SQLString = SQLString & " order by d.DOD_DISP_SEQ"
            REM **********************
            do_dt = gDB.getDataTable(SQLString)
            For i = 0 To do_dt.Rows.Count - 1
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
                dt.Rows(rows_count - 1).Item("SD_SEQ") = Session("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("SD_DISP_SEQ") = do_dt.Rows(i).Item("DOD_DISP_SEQ")
                dt.Rows(rows_count - 1).Item("SD_PALLET_NO") = do_dt.Rows(i).Item("DOD_PALLET_NO")
                dt.Rows(rows_count - 1).Item("SD_CARTON_NO") = do_dt.Rows(i).Item("DOD_CARTON_NO")
                dt.Rows(rows_count - 1).Item("SD_ITM_CODE") = do_dt.Rows(i).Item("DOD_ITM_CODE")
                dt.Rows(rows_count - 1).Item("SD_PACK_KEY") = do_dt.Rows(i).Item("DOD_PACK_KEY")
                dt.Rows(rows_count - 1).Item("SD_ITM_DESC") = do_dt.Rows(i).Item("DOD_ITM_DESC")
                dt.Rows(rows_count - 1).Item("SD_PACK_NO") = do_dt.Rows(i).Item("DOD_DISP_SEQ")
                dt.Rows(rows_count - 1).Item("SD_PACK_TYPE") = do_dt.Rows(i).Item("DOD_PACK_TYPE")
                dt.Rows(rows_count - 1).Item("SD_QTY") = do_dt.Rows(i).Item("DOD_QTY")
                dt.Rows(rows_count - 1).Item("SD_UOM") = do_dt.Rows(i).Item("DOD_UOM")
                dt.Rows(rows_count - 1).Item("SD_PCS_UOM") = do_dt.Rows(i).Item("DOD_PCS_UOM")
                dt.Rows(rows_count - 1).Item("SD_TOTPCS") = do_dt.Rows(i).Item("DOD_TOTPCS")
                dt.Rows(rows_count - 1).Item("SD_TOT_WGT") = do_dt.Rows(i).Item("DOD_TOT_WGT")
                dt.Rows(rows_count - 1).Item("SD_TOT_CBM") = do_dt.Rows(i).Item("DOD_TOT_CBM")
                dt.Rows(rows_count - 1).Item("SD_REM") = do_dt.Rows(i).Item("DOD_REM")
                dt.Rows(rows_count - 1).Item("SD_TRACK_NO") = do_dt.Rows(i).Item("DOD_TRACK_NO")

                dt.Rows(rows_count - 1).Item("DO_CODE") = do_dt.Rows(i).Item("DO_CODE")
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
            Next
            dt.AcceptChanges()
            Session("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here

        Call cU.newChangeGVLabel(GridView1, "No.", "編號")
        Call cU.newChangeGVLabel(GridView1, "Seq.", "序號")
        Call cU.newChangeGVLabel(GridView1, "Tracking No.", "追查編號")
        Call cU.newChangeGVLabel(GridView1, "Pallet No.", "工作編號")
        Call cU.newChangeGVLabel(GridView1, "Cartono. No.", "参考編號")
        Call cU.newChangeGVLabel(GridView1, "Item Code.", "物件號碼")
        Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        Call cU.newChangeGVLabel(GridView1, "Pack No.", "封裝形式")
        Call cU.newChangeGVLabel(GridView1, "Pack Type", "封裝形式")
        Call cU.newChangeGVLabel(GridView1, "Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        Call cU.newChangeGVLabel(GridView1, "PCS/ UOM", "單位件數")
        Call cU.newChangeGVLabel(GridView1, "Total PCS", "總件數")
        Call cU.newChangeGVLabel(GridView1, "Total Wgt", "總重量")
        Call cU.newChangeGVLabel(GridView1, "Total CBM", "總體積(cbm)")
        Call cU.newChangeGVLabel(GridView1, "Remarks", "備註")
        Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Protected Sub btnShip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShip.Click
        STATUS.Text = "SHIPPED"
        STATUS.ForeColor = Drawing.Color.Blue
        btnShip.Enabled = False
        ar.sec_viewMode = "Y"
        Call save()
        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
    End Sub
End Class
