Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing


Partial Class MASTER_IM_altVendorMaster
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As New AccessRightUtils
    Private thumb As ThumbGenerator

    Private dt As New DataTable

    Private imp_code As String = ""

    Private SYSP_TEMP_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM ****************************
        REM Modify Access Right Here

        ar = New AccessRightUtils("MAST_IM", Session("usr_id"), Me)
        ar.hideForm(Me)

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("mypagemode") = Nothing
            Session("mypagemode") = Request("r")

            ViewState("remove_image1") = False
            ViewState("remove_image2") = False

            'uiFun.load_dropdown(vnd_code0, "select VND_CODE, VND_NAME from WMS_VENDOR ORDER BY 2", "VND_CODE", "VND_NAME", , Session("gSelectLabel"))
            Session.Remove(aitm_picture1_upload.ClientID)
            Session.Remove(aitm_picture2_upload.ClientID)

            uiFun.load_dropdown(aitm_pref_wh, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
            uiFun.load_dropdown(aitm_uom, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
        Else
            If aitm_picture1_upload.HasFile Then
                Session(aitm_picture1_upload.ClientID) = aitm_picture1_upload.PostedFile
                aitm_picture1_upload_lit.Text = aitm_picture1_upload.PostedFile.FileName
                aitm_picture1_upload.Visible = False
                aitm_picture1_upload_lit.Visible = True
                aitm_picture1_edit.Visible = True
            End If

            If aitm_picture2_upload.HasFile Then
                Session(aitm_picture2_upload.ClientID) = aitm_picture2_upload.PostedFile
                aitm_picture2_upload_lit.Text = aitm_picture2_upload.PostedFile.FileName
                aitm_picture2_upload.Visible = False
                aitm_picture2_upload_lit.Visible = True
                aitm_picture2_edit.Visible = True
            End If

        End If

        If Session("mypagemode") = "N" Then
            itm_code.Text = Session("im_itm_code")

            aitm_ref_price.ReadOnly = False
            aitm_ref_curr.ReadOnly = False
        Else
            aitm_ref_price.BackColor = Drawing.Color.Transparent
            aitm_ref_price.BorderStyle = BorderStyle.None
            aitm_ref_curr.BackColor = Drawing.Color.Transparent
            aitm_ref_curr.BorderStyle = BorderStyle.None
            aitm_ref_price.ReadOnly = True
            aitm_ref_curr.ReadOnly = True
        End If
        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Alt. Vendors Item"
            lbl_vnd_code.Text = "Vendor:"
            lbl_itm_code.Text = "Item Code:"
            lbl_aitm_status.Text = "Status:"
            lbl_alv_date_added.Text = "Vendor Added On:"
            lbl_alv_vnd_itmcode.Text = "Vendor's Item Code:"
            lbl_pack_key.Text = "Pack Key:"
            lbl_storer_code.Text = "Storer:"
            lbl_aitm_name.Text = "Item Name:"
            lbl_aitm_uom.Text = "Unit of Measure:"
            lbl_aitm_desc.Text = "Item Description:"
            lbl_aitm_spec.Text = "Item Specifications:"
            lbl_aitm_remarks.Text = "Item Remarks:"
            lbl_aitm_pcs_per_pack.Text = "Number Per UOM:"
            lbl_aitm_qty_per_ctn.Text = "Qty Per Carton:"
            lbl_sizes.Text = "Carton Size (CM):"
            lbl_aitm_length.Text = "Length"
            lbl_aitm_width.Text = "Width"
            lbl_aitm_hight.Text = "Height"
            lbl_aitm_cbm.Text = "Vol (CBM):"
            lbl_aitm_vol.Text = "Weight (KG):"
            lbl_aitm_pref_wh.Text = "Preferred Warehouse:"
            lbl_aitm_pref_loc.Text = "Preferred Location:"
            lbl_aitm_ref_curr.Text = "Ref. Currency:"
            lbl_aitm_ref_price.Text = "Ref. Price:"
            lbl_AITM_ORIGIN.Text = "Origin:"
            lbl_AITM_GROSS_WEIGHT.Text = "G.W.:"
            lbl_AITM_NET_WEIGHT.Text = "N.W.:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"

            lbl_ImageHd.Text = "Images"

            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this alt. vendor item?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this alt. vendor item?"");"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "Alt. Vendors Item"
            lbl_vnd_code.Text = "Vendor:"
            lbl_itm_code.Text = "Item Code:"
            lbl_aitm_status.Text = "Status:"
            lbl_alv_date_added.Text = "Vendor Added On:"
            lbl_alv_vnd_itmcode.Text = "Vendor's Item Code:"
            lbl_pack_key.Text = "Pack Key:"
            lbl_storer_code.Text = "Storer:"
            lbl_aitm_name.Text = "Item Name:"
            lbl_aitm_uom.Text = "Unit of Measure:"
            lbl_aitm_desc.Text = "Item Description:"
            lbl_aitm_spec.Text = "Item Specifications:"
            lbl_aitm_remarks.Text = "Item Remarks:"
            lbl_aitm_pcs_per_pack.Text = "Number Per UOM:"
            lbl_aitm_qty_per_ctn.Text = "Qty Per Carton:"
            lbl_sizes.Text = "Carton Size (CM):"
            lbl_aitm_length.Text = "Length"
            lbl_aitm_width.Text = "Width"
            lbl_aitm_hight.Text = "Height"
            lbl_aitm_cbm.Text = "Vol (CBM):"
            lbl_aitm_vol.Text = "Weight (KG):"
            lbl_aitm_pref_wh.Text = "Preferred Warehouse:"
            lbl_aitm_pref_loc.Text = "Preferred Location:"
            lbl_aitm_ref_curr.Text = "Ref. Currency:"
            lbl_aitm_ref_price.Text = "Ref. Price:"
            lbl_AITM_ORIGIN.Text = "Origin:"
            lbl_AITM_GROSS_WEIGHT.Text = "G.W.:"
            lbl_AITM_NET_WEIGHT.Text = "N.W.:"

            lbl_sys_cb.Text = "创建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "创建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"

            lbl_ImageHd.Text = "Images"

            saveBtn1.OnClientClick = "return confirm(""确定保存资料?"");"
            saveBtn2.OnClientClick = "return confirm(""确定保存资料?"");"

        End If
        REM **********************

        REM **********************
        REM Additional CSS
        'pack_key.CssClass = "REQUIRED"
        'storer_code.CssClass = "REQUIRED"


        REM **********************

        If Not IsPostBack Then
            ViewState("avi_dt") = Nothing
            ViewState("image1_name") = ""
            ViewState("image2_name") = ""

            ViewState("remove_image1") = False
            ViewState("remove_image2") = False

            If Not Session("avi_dt") Is Nothing Then
                ViewState("avi_dt") = Session("avi_dt")
                dt = Session("avi_dt")
            End If

            If Session("mypagemode") = "N" Then
                itm_code.Text = Session("im_itm_code")
                pack_key.Text = Session("im_pack_key")
                storer_code.Text = Session("im_storer_code")

                aitm_picture1.Enabled = False
                aitm_picture2.Enabled = False
                aitm_picture1.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                aitm_picture2.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
            Else
                If dt.Rows.Count > 0 Then
                    Dim vndCode As String = ""
                    Dim pic_preview_yn As Boolean = False

                    If Request("vnd_code") IsNot Nothing Then vndCode = Request("vnd_code")

                    Dim vndRow As DataRow() = dt.Select("vnd_code = '" & vndCode & "'")

                    If vndRow.Count > 0 Then
                        itm_code.Text = vndRow(0).Item("itm_code").ToString
                        pack_key.Text = vndRow(0).Item("pack_key").ToString
                        aitm_status.Text = gU.decodeNullOrEmpty(vndRow(0).Item("aitm_status").ToString, "Active")

                        storer_code.Text = vndRow(0).Item("storer_code").ToString
                        alv_date_added.Text = vndRow(0).Item("alv_date_added").ToString

                        vnd_code.Text = vndRow(0).Item("vnd_code").ToString
                        h_vnd_code.Value = vndRow(0).Item("vnd_code").ToString
                        vnd_name.Text = vndRow(0).Item("vnd_name").ToString
                        h_vnd_name.Value = vndRow(0).Item("vnd_code").ToString

                        alv_vnd_itmcode.Text = vndRow(0).Item("alv_vnd_itmcode").ToString

                        aitm_name.Text = vndRow(0).Item("aitm_name").ToString
                        aitm_uom.SelectedValue = vndRow(0).Item("aitm_uom").ToString

                        aitm_desc.Text = vndRow(0).Item("aitm_desc").ToString
                        aitm_spec.Text = vndRow(0).Item("aitm_spec").ToString
                        aitm_remarks.Text = vndRow(0).Item("aitm_remarks").ToString

                        aitm_pcs_per_pack.Text = vndRow(0).Item("aitm_pcs_per_pack").ToString

                        aitm_qty_per_ctn.Text = vndRow(0).Item("aitm_qty_per_ctn").ToString

                        aitm_length.Text = vndRow(0).Item("aitm_length").ToString
                        aitm_width.Text = vndRow(0).Item("aitm_width").ToString
                        aitm_hight.Text = vndRow(0).Item("aitm_hight").ToString

                        aitm_cbm.Text = vndRow(0).Item("aitm_cbm").ToString
                        aitm_vol.Text = vndRow(0).Item("aitm_vol").ToString

                        aitm_pref_wh.SelectedValue = vndRow(0).Item("aitm_pref_wh").ToString
                        aitm_pref_loc.Text = vndRow(0).Item("aitm_pref_loc").ToString

                        aitm_ref_curr.Text = vndRow(0).Item("aitm_ref_curr").ToString
                        aitm_ref_price.Text = vndRow(0).Item("aitm_ref_price").ToString

                        AITM_ORIGIN.Text = vndRow(0).Item("AITM_ORIGIN").ToString
                        AITM_GROSS_WEIGHT.Text = vndRow(0).Item("AITM_GROSS_WEIGHT").ToString
                        AITM_NET_WEIGHT.Text = vndRow(0).Item("AITM_NET_WEIGHT").ToString

                        ViewState("image1_name") = vndRow(0).Item("aitm_picture1").ToString
                        ViewState("image2_name") = vndRow(0).Item("aitm_picture2").ToString

                        aitm_picture1.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & vndRow(0).Item("aitm_picture1").ToString & "&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                        aitm_picture2.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & vndRow(0).Item("aitm_picture2").ToString & "&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")

                        sys_cb.Text = vndRow(0).Item("sys_cb").ToString
                        sys_lub.Text = vndRow(0).Item("sys_lub").ToString
                        sys_cd.Text = cU.chgToFullDF(vndRow(0).Item("sys_cd").ToString)
                        sys_lud.Text = cU.chgToFullDF(vndRow(0).Item("sys_lud").ToString)

                        If vndRow(0).Item("aitm_picture1").ToString <> "" Then
                            aitm_picture1.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & vndRow(0).Item("aitm_picture1").ToString & "&refresh=true&ds=true&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                            aitm_picture1.Target = "_new1"

                            preview_pic1.Visible = True
                            aitm_picture1_remove.Visible = True
                            aitm_picture1.Enabled = True
                            pic_preview_yn = True
                        Else
                            preview_pic1.Visible = False
                            aitm_picture1_remove.Visible = False
                            aitm_picture1.Enabled = False
                        End If

                        If vndRow(0).Item("aitm_picture2").ToString <> "" Then
                            aitm_picture2.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & vndRow(0).Item("aitm_picture2").ToString & "&refresh=true&ds=true&Height=180&width=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                            aitm_picture2.Target = "_new2"

                            preview_pic2.Visible = True
                            aitm_picture2_remove.Visible = True
                            aitm_picture2.Enabled = True
                            pic_preview_yn = True
                        Else
                            preview_pic2.Visible = False
                            aitm_picture2_remove.Visible = False
                            aitm_picture2.Enabled = False
                        End If

                        If pic_preview_yn Then
                            pic_tr.Visible = True
                            remove_tr.Visible = True
                        Else
                            pic_tr.Visible = False
                            remove_tr.Visible = False
                        End If
                    End If              
                End If
            End If
        Else
            dt = ViewState("avi_dt")
        End If

        If aitm_status.Text = "CANCELLED" Then
            ar.sec_write = "N"
        End If
        ' "<%=h_vnd_code.clientID %>" + "|, " + "<%=vnd_code.ClientID %>" + "|, " + "<%=h_vnd_name.ClientID %>" + "|1, "  + "<%=vnd_name.ClientID %>" + "|1");
        btnVnd_LookUp.Attributes.Add("onclick", "javascript:VendorLookUp('" & h_vnd_code.ClientID & "','" & vnd_code.ClientID & "','" & h_vnd_name.ClientID & "','" & vnd_name.ClientID & "')")

        'ar.setFieldCustomize(Me, "MAST_IM", storer_code.Text, "WMS_ALT_VEND_ITEM", "D")

        vnd_code.Attributes.Add("readonly", "readonly")

    End Sub

    Private Function validateAll() As Boolean

        If h_vnd_code.Value.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_vnd_code.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_vnd_code.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If alv_date_added.Text.Trim <> "" And Not gU.isValidDate(alv_date_added.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_alv_date_added.Text & "Invalid Date!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_alv_date_added.Text & "無效的日期!", Session("gLang"))
            End If
            Return False
        End If

        If aitm_pcs_per_pack.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_aitm_pcs_per_pack.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_aitm_pcs_per_pack.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""

        Dim tmp_itm_flags As String = ""
        Dim nFileName1 As String = ""
        Dim nFileName2 As String = ""
        Dim tmpImpCode As String = ""
        Dim tempRemoveFile1 As Boolean = False
        Dim tempRemoveFile2 As Boolean = False

        Dim aUP1 As HttpPostedFile = Nothing
        Dim aUP2 As HttpPostedFile = Nothing

        If Not Session(aitm_picture1_upload.ClientID) Is Nothing Then aUP1 = TryCast(Session(aitm_picture1_upload.ClientID), HttpPostedFile)
        If Not Session(aitm_picture2_upload.ClientID) Is Nothing Then aUP2 = TryCast(Session(aitm_picture2_upload.ClientID), HttpPostedFile)

        If imp_code = "" Then tmpImpCode = "W" Else tmpImpCode = imp_code

        Dim old_filename1 As String = ViewState("image1_name")
        Dim old_filename2 As String = ViewState("image2_name")

        Dim tmpFInfo As IO.FileInfo

        If validateAll() Then
            Try
                Dim RowIndex As Integer = 0
                Dim ifDateValid As Boolean = False
                Dim tempDate As DateTime

                If Session("mypagemode") = "N" Then

                    If aUP1 Is Nothing Then
                        If aitm_picture1_upload.HasFile Then
                            tmpFInfo = New FileInfo(aitm_picture1_upload.PostedFile.FileName)
                            nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_1_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(aUP1.FileName)
                        nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_1_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                    End If

                    If aUP2 Is Nothing Then
                        If aitm_picture2_upload.HasFile Then
                            tmpFInfo = New FileInfo(aitm_picture2_upload.PostedFile.FileName)
                            nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_2_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(aUP2.FileName)
                        nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_2_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                    End If

                    dt.Rows.Add()

                    RowIndex = dt.Rows.Count - 1

                    dt.Rows(RowIndex).Item("itm_code") = itm_code.Text
                    dt.Rows(RowIndex).Item("pack_key") = pack_key.Text
                    dt.Rows(RowIndex).Item("aitm_status") = aitm_status.Text
                    dt.Rows(RowIndex).Item("storer_code") = storer_code.Text

                    'tempDate = gU.convertToDate(alv_date_added.Text, ifDateValid)

                    'If ifDateValid Then
                    '    dt.Rows(RowIndex).Item("alv_date_added") = tempDate
                    'Else
                    '    dt.Rows(RowIndex).Item("alv_date_added") = DBNull.Value
                    'End If

                    dt.Rows(RowIndex).Item("alv_date_added") = alv_date_added.Text

                    'dt.Rows(RowIndex).Item("vnd_code") = vnd_code.Text
                    'dt.Rows(RowIndex).Item("vnd_name") = vnd_name.Text

                    dt.Rows(RowIndex).Item("vnd_code") = h_vnd_code.Value
                    dt.Rows(RowIndex).Item("vnd_name") = h_vnd_name.Value
                    dt.Rows(RowIndex).Item("aitm_status") = "Active"
                    dt.Rows(RowIndex).Item("alv_vnd_itmcode") = alv_vnd_itmcode.Text
                    dt.Rows(RowIndex).Item("aitm_name") = aitm_name.Text
                    dt.Rows(RowIndex).Item("aitm_uom") = aitm_uom.SelectedValue
                    dt.Rows(RowIndex).Item("aitm_desc") = aitm_desc.Text
                    dt.Rows(RowIndex).Item("aitm_spec") = aitm_spec.Text
                    dt.Rows(RowIndex).Item("aitm_remarks") = aitm_remarks.Text
                    dt.Rows(RowIndex).Item("aitm_pcs_per_pack") = gU.decodeEmptyCdbl(aitm_pcs_per_pack.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_qty_per_ctn") = gU.decodeEmptyCdbl(aitm_qty_per_ctn.Text, 1)
                    dt.Rows(RowIndex).Item("aitm_length") = gU.decodeEmptyCdbl(aitm_length.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_width") = gU.decodeEmptyCdbl(aitm_width.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_hight") = gU.decodeEmptyCdbl(aitm_hight.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_cbm") = gU.decodeEmptyCdbl(aitm_cbm.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_vol") = gU.decodeEmptyCdbl(aitm_vol.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_pref_wh") = aitm_pref_wh.SelectedValue
                    dt.Rows(RowIndex).Item("aitm_pref_loc") = aitm_pref_loc.Text
                    dt.Rows(RowIndex).Item("AITM_ORIGIN") = AITM_ORIGIN.Text
                    dt.Rows(RowIndex).Item("AITM_GROSS_WEIGHT") = gU.decodeEmptyCdbl(AITM_GROSS_WEIGHT.Text, 0)
                    dt.Rows(RowIndex).Item("AITM_NET_WEIGHT") = gU.decodeEmptyCdbl(AITM_NET_WEIGHT.Text, 0)

                    dt.Rows(RowIndex).Item("aitm_ref_curr") = aitm_ref_curr.Text
                    dt.Rows(RowIndex).Item("aitm_ref_price") = gU.decodeEmptyCdbl(aitm_ref_price.Text, 0)

                    dt.Rows(RowIndex).Item("aitm_picture1") = nFileName1
                    dt.Rows(RowIndex).Item("aitm_picture2") = nFileName2

                    dt.Rows(RowIndex).Item("mFlag") = "N"

                    dt.AcceptChanges()

                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    RowIndex = CInt(Request("r").ToString)

                    If aUP1 Is Nothing Then
                        If aitm_picture1_upload.HasFile Then
                            tmpFInfo = New FileInfo(aitm_picture1_upload.PostedFile.FileName)
                            nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_1_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(aUP1.FileName)
                        nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_1_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                    End If

                    If aUP2 Is Nothing Then
                        If aitm_picture2_upload.HasFile Then
                            tmpFInfo = New FileInfo(aitm_picture2_upload.PostedFile.FileName)
                            nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_2_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(aUP2.FileName)
                        nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_ALTV_2_" & gU.decodeNullOrEmpty(vnd_code.Text, "0") & gU.decodeNullOrEmpty(itm_code.Text, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.Text, "0") & tmpFInfo.Extension
                    End If

                    If Not ViewState("remove_image1") Is Nothing Then
                        tempRemoveFile1 = ViewState("remove_image1")
                    Else
                        tempRemoveFile1 = False
                    End If

                    If Not ViewState("remove_image2") Is Nothing Then
                        tempRemoveFile2 = ViewState("remove_image2")
                    Else
                        tempRemoveFile2 = False
                    End If

                    dt.Rows(RowIndex).Item("itm_code") = itm_code.Text
                    dt.Rows(RowIndex).Item("pack_key") = pack_key.Text
                    dt.Rows(RowIndex).Item("aitm_status") = aitm_status.Text
                    dt.Rows(RowIndex).Item("storer_code") = storer_code.Text

                    'tempDate = gU.convertToDate(alv_date_added.Text, ifDateValid)

                    'If ifDateValid Then
                    '    dt.Rows(RowIndex).Item("alv_date_added") = tempDate
                    'Else
                    '    dt.Rows(RowIndex).Item("alv_date_added") = DBNull.Value
                    'End If

                    dt.Rows(RowIndex).Item("alv_date_added") = alv_date_added.Text

                    'dt.Rows(RowIndex).Item("vnd_code") = vnd_code.Text
                    'dt.Rows(RowIndex).Item("vnd_name") = vnd_name.Text
                    dt.Rows(RowIndex).Item("vnd_code") = h_vnd_code.Value
                    dt.Rows(RowIndex).Item("vnd_name") = h_vnd_name.Value
                    dt.Rows(RowIndex).Item("alv_vnd_itmcode") = alv_vnd_itmcode.Text
                    dt.Rows(RowIndex).Item("aitm_name") = aitm_name.Text
                    dt.Rows(RowIndex).Item("aitm_uom") = aitm_uom.SelectedValue
                    dt.Rows(RowIndex).Item("aitm_desc") = aitm_desc.Text
                    dt.Rows(RowIndex).Item("aitm_spec") = aitm_spec.Text
                    dt.Rows(RowIndex).Item("aitm_remarks") = aitm_remarks.Text
                    dt.Rows(RowIndex).Item("aitm_pcs_per_pack") = gU.decodeEmptyCdbl(aitm_pcs_per_pack.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_qty_per_ctn") = gU.decodeEmptyCdbl(aitm_qty_per_ctn.Text, 1)
                    dt.Rows(RowIndex).Item("aitm_length") = gU.decodeEmptyCdbl(aitm_length.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_width") = gU.decodeEmptyCdbl(aitm_width.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_hight") = gU.decodeEmptyCdbl(aitm_hight.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_cbm") = gU.decodeEmptyCdbl(aitm_cbm.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_vol") = gU.decodeEmptyCdbl(aitm_vol.Text, 0)
                    dt.Rows(RowIndex).Item("aitm_pref_wh") = aitm_pref_wh.SelectedValue
                    dt.Rows(RowIndex).Item("aitm_pref_loc") = aitm_pref_loc.Text

                    dt.Rows(RowIndex).Item("AITM_ORIGIN") = AITM_ORIGIN.Text
                    dt.Rows(RowIndex).Item("AITM_GROSS_WEIGHT") = gU.decodeEmptyCdbl(AITM_GROSS_WEIGHT.Text, 0)
                    dt.Rows(RowIndex).Item("AITM_NET_WEIGHT") = gU.decodeEmptyCdbl(AITM_NET_WEIGHT.Text, 0)


                    dt.Rows(RowIndex).Item("aitm_ref_curr") = aitm_ref_curr.Text
                    dt.Rows(RowIndex).Item("aitm_ref_price") = gU.decodeEmptyCdbl(aitm_ref_price.Text, 0)

                    REM **********************
                End If

                If Not IO.Directory.Exists(SYSP_TEMP_DIR) Then IO.Directory.CreateDirectory(SYSP_TEMP_DIR)

                Dim isSavePic1 As Boolean = False
                Dim isSavePic2 As Boolean = False

                If Not tempRemoveFile1 Then
                    If aUP1 Is Nothing Then
                        If aitm_picture1_upload.HasFile Then
                            aitm_picture1_upload.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & nFileName1)
                            isSavePic1 = True
                        End If
                    Else
                        aUP1.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & nFileName1)
                        isSavePic1 = True
                    End If

                    If isSavePic1 Then
                        dt.Rows(RowIndex).Item("aitm_picture1") = nFileName1
                        dt.Rows(RowIndex).Item("has_pic1") = "Y"
                    End If
                Else
                    ViewState("remove_image1") = False
                    dt.Rows(RowIndex).Item("remove_pic1") = "Y"
                    dt.Rows(RowIndex).Item("aitm_picture1") = DBNull.Value
                End If

                If Not tempRemoveFile2 Then
                    If aUP2 Is Nothing Then
                        If aitm_picture2_upload.HasFile Then
                            aitm_picture2_upload.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & nFileName2)
                            isSavePic2 = True
                        End If
                    Else
                        aUP2.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & nFileName2)
                        isSavePic2 = True
                    End If

                    If isSavePic2 Then
                        dt.Rows(RowIndex).Item("aitm_picture2") = nFileName2
                        dt.Rows(RowIndex).Item("has_pic2") = "Y"
                    End If
                Else
                    ViewState("remove_image2") = False
                    dt.Rows(RowIndex).Item("remove_pic2") = "Y"
                    dt.Rows(RowIndex).Item("aitm_picture2") = DBNull.Value
                End If

                dt.AcceptChanges()

                If Session("mypagemode") = "N" Then Session.Remove("mypagemode")

                toBringOver(dt)

                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", _
                "<script language=""JavaScript"">window.open('','_self');window.close();</script>")

            Catch ex As Exception

                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Private Sub toBringOver(ByRef fromTable As DataTable)
        Dim scriptString As String

        Session("avi_dt") = fromTable
        Session.Remove(aitm_picture1_upload.ClientID)
        Session.Remove(aitm_picture2_upload.ClientID)

        scriptString = "<script language=""JavaScript""> window.opener.document.forms(0).submit(); </script>"

        If Not Page.ClientScript.IsClientScriptBlockRegistered(scriptString) Then
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "script", scriptString)
        End If

    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub aitm_picture1_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles aitm_picture1_remove.Click
        ViewState("remove_image1") = True

        aitm_picture1_upload.Enabled = False
        aitm_picture1_remove.Enabled = False
        aitm_picture1.Enabled = False
        preview_pic1.Enabled = False

    End Sub

    Protected Sub aitm_picture2_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles aitm_picture2_remove.Click
        ViewState("remove_image2") = True

        aitm_picture2_upload.Enabled = False
        aitm_picture2_remove.Enabled = False
        aitm_picture2.Enabled = False
        preview_pic2.Enabled = False

    End Sub
End Class
