Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_DELI_ORDER_DeliOrderNote
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private uiFun As New UIfunc
    Private DB As New DBfunc


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OB_MDN", Session("usr_id"), Me)

        ar.hideForm(Me)

        If Not IsPostBack Then
            Session("nDataSource") = ""

            Select Case Session("gLang")
                Case "C"
                    lHeader.Text = "批次列印送貨單"
                    btnGen.Text = "列印"
                    lbl_STORER_CODE.Text = "貨主號碼"
                    lbl_EDI_SIR_NO.Text = "客戶訂單號碼"
                    lbl_WP_NO.Text = "波次號"
                Case Else
                    lHeader.Text = "Batch Print Delivery Note:"
                    btnGen.Text = "Print"
                    lbl_STORER_CODE.Text = "Storer Code"
                    lbl_EDI_SIR_NO.Text = "Customer Order No."
                    lbl_WP_NO.Text = "Wave Picking No."
            End Select

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))

        End If

    End Sub

    Protected Sub btnGen_Click(sender As Object, e As System.EventArgs) Handles btnGen.Click
        If Not String.IsNullOrWhiteSpace(EDI_SIR_NO.Text) Or Not String.IsNullOrWhiteSpace(WP_NO.Text) Then

            Dim sqlString As String
            Dim nDataSource As DataTable
            Dim Paras(1) As ReportParameter

            Dim lEDI As String = ""
            Dim lstorer_code As String = ""
            Dim lWP_NO As String = ""
            Dim tempStr As String = ""



            If STORER_CODE.SelectedValue <> "" Then
                lstorer_code = STORER_CODE.SelectedValue

                tempStr &= "and dm.storer_code='" & gU.dbEncode(lstorer_code) & "' "
            End If

            If Not String.IsNullOrWhiteSpace(WP_NO.Text) Then
                lWP_NO = gU.dbConvList(WP_NO.Text.Trim)

                tempStr &= " AND EXISTS (SELECT 1 FROM WMS_DO_PICKLIST_D WHERE WMS_DO_PICKLIST_D.IMP_CODE = dm.IMP_CODE AND WMS_DO_PICKLIST_D.STORER_CODE=dm.STORER_CODE AND WMS_DO_PICKLIST_D.DO_CODE = dm.DO_CODE " & _
                          " AND PLD_WAVE_PICK_NO IN (" & lWP_NO & "))"
            End If

            If Not String.IsNullOrWhiteSpace(EDI_SIR_NO.Text) Then
                lEDI = gU.dbConvList(EDI_SIR_NO.Text.Trim)
                tempStr &= "and dm.DO_EDI_SIR_NO IN (" & lEDI & ") "
            End If

            sqlString = " SELECT dm.STORER_CODE, dm.DO_CODE, dm.DO_ISSUED_BY, convert(varchar, dm.DO_DATE, 103) AS do_date, " & _
                     " Convert(varchar, dm.DO_CONF_DELDATE, 103) AS do_conf_deldate, Convert(varchar(5), dm.DO_CONF_DELTIME, 108) AS do_conf_deltime, dm.CUS_NAME, " & _
                     " (CASE WHEN not dm.do_addr1 is null THEN dm.do_addr1 ELSE '' END + CASE WHEN not dm.do_addr2 is null THEN char(13) + char(10) " & _
                     " + dm.do_addr2 ELSE '' END + CASE WHEN not dm.do_addr3 is null THEN char(13) + char(10) " & _
                     " + dm.do_addr3 ELSE '' END + CASE WHEN not dm.do_area_del is null THEN char(13) + char(10) " & _
                     " + dm.do_area_del ELSE '' END + CASE WHEN not dm.do_region_del is null  " & _
                     " THEN ' ' + dm.do_region_del ELSE '' END + CASE WHEN not dm.do_country_del is null THEN char(13) + char(10) " & _
                     " + dm.do_country_del ELSE '' END) AS do_addr, dm.DO_CUS_CONT, dm.DO_CUS_CONT_TEL, dm.DO_CUS_REF_NO, dd.DOD_DISP_SEQ, " & _
                     " dd.DOD_ITM_CODE, itm.ITM_DESC as DOD_ITM_DESC, dd.DOD_QTY, dd.DOD_REM, convert(decimal, ISNULL(dm.DO_TOTL_PALLETS, 0)) AS total_pallet, " & _
                     " convert(decimal, ISNULL(dm.DO_TOTL_CARTONS, 0)) AS total_carton, dm.DO_DRIVER, dm.DO_INV_NO,dd.DOD_BATCH_NO, " & _
                     " WMS_VENDOR.VND_NAME,itm.ITM_SKU_NO,dd.DOD_TOT_CBM, dd.DOD_TOT_WGT, Convert(varchar, DOD_EXPIRY_DATE,103) as DOD_EXPIRY_DATE, Convert(varchar, DOD_MANU_DATE,103) as DOD_MANU_DATE, " & _
                     " DO_CONSIGNEE, (CASE WHEN not dm.DO_CONSIGNEE_ADDR1 is null THEN dm.DO_CONSIGNEE_ADDR1 ELSE '' END + CASE WHEN not dm.DO_CONSIGNEE_ADDR2 is null THEN char(13) + char(10) " & _
                     " + dm.DO_CONSIGNEE_ADDR2 ELSE '' END + CASE WHEN not dm.DO_CONSIGNEE_ADDR3 is null THEN char(13) + char(10) " & _
                     " + dm.DO_CONSIGNEE_ADDR3 ELSE '' END) as DO_CONSIGNEE_ADDR1, itm.ITM_MODEL, dm.DO_FTRACK_NO,DO_DELIVERY_RMKS,Convert(varchar, DO_ARRIVAL_DATE,103) as DO_ARRIVAL_DATE, DO_POST_CODE,DO_REM,DO_EDI_WIT_NO,dm.DO_EDI_SIR_NO, " & _
                     " CONVERT(varchar, DO_TARGET_DELDATE, 103) as DO_TARGET_DELDATE, '' as do_code_img, '' as ftrack_img, '' as PLD_WAVE_PICK_NO " & _
                     " FROM WMS_DELV_ORDER dm INNER JOIN " & _
                     " WMS_DELV_ORDER_D dd ON dm.IMP_CODE = dd.IMP_CODE AND dm.STORER_CODE = dd.STORER_CODE AND " & _
                     " dm.DO_CODE = dd.DO_CODE Left outer join " & _
                     " WMS_VENDOR ON dd.IMP_CODE = WMS_VENDOR.IMP_CODE AND dd.STORER_CODE = WMS_VENDOR.STORER_CODE AND " & _
                     " dd.DOD_VND_CODE = WMS_VENDOR.VND_CODE " & _
                     " LEFT OUTER JOIN WMS_ITEM itm " & _
                     " ON dd.IMP_CODE = itm.IMP_CODE AND dd.STORER_CODE=itm.STORER_CODE AND dd.DOD_ITM_CODE = itm.ITM_CODE AND dd.DOD_PACK_KEY = itm.PACK_KEY " & _
                    "Where dm.imp_code = '" & Session("imp_code") & "' " & _
                    tempStr & _
                    "order by dm.DO_CODE, dd.dod_disp_seq "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then
                Dim oBmp As Bitmap
                Dim iBCWriter As New ZXing.BarcodeWriter

                Dim ImgByte() As Byte
                Dim ByteString As String

                Dim ldo_code As String
                Dim PreDO_CODE As String = ""

                Dim Mwp_no As String = ""
                Dim fTrackNo As String = ""

                For i = 0 To nDataSource.Rows.Count - 1
                    If PreDO_CODE <> nDataSource.Rows(i).Item("DO_CODE").ToString.Trim AndAlso nDataSource.Rows(i).Item("DO_CODE").ToString.Trim <> "" Then
                        ldo_code = gU.decodeNullOrEmpty(nDataSource.Rows(i).Item("DO_CODE").ToString.Trim, "-")

                        iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE
                        iBCWriter.Options.PureBarcode = True
                        oBmp = iBCWriter.Write(ldo_code)

                        Using IOStream As New MemoryStream
                            oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                            ImgByte = IOStream.ToArray

                            ByteString = Convert.ToBase64String(ImgByte)
                            nDataSource.Rows(i).Item("do_code_img") = ByteString
                        End Using

                        iBCWriter.Format = ZXing.BarcodeFormat.CODE_128

                        fTrackNo = DB.getValueFromSQL("SELECT MAX(DOB_FTRACK_NO) as DOB_FTRACK_NO from WMS_DO_BOX_TRACK WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(nDataSource.Rows(i).Item("STORER_CODE").ToString.Trim) & "' AND DO_CODE='" & gU.dbEncode(nDataSource.Rows(i).Item("DO_CODE").ToString.Trim) & "'")
                        oBmp = iBCWriter.Write(gU.decodeNullOrEmpty(fTrackNo, "-"))
                        nDataSource.Rows(i).Item("DO_FTRACK_NO") = fTrackNo

                        Using IOStream As New MemoryStream
                            oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                            ImgByte = IOStream.ToArray

                            ByteString = Convert.ToBase64String(ImgByte)

                            nDataSource.Rows(i).Item("ftrack_img") = ByteString
                        End Using

                        Mwp_no = DB.getValueFromSQL("SELECT MAX(PLD_WAVE_PICK_NO) as PLD_WAVE_PICK_NO from WMS_DO_PICKLIST_D WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(nDataSource.Rows(i).Item("STORER_CODE").ToString.Trim) & "' AND DO_CODE='" & gU.dbEncode(nDataSource.Rows(i).Item("DO_CODE").ToString.Trim) & "'")
                        nDataSource.Rows(i).Item("PLD_WAVE_PICK_NO") = Mwp_no


                        PreDO_CODE = nDataSource.Rows(i).Item("DO_CODE").ToString.Trim
                    End If
                Next

                nDataSource.AcceptChanges()

                Session("nDataSource") = nDataSource
                iframe.Attributes.Add("src", "downloadNote.aspx")
            Else
                Dim alertMsg As String = ""
                Select Case Session("gLang")
                    Case "C"
                        alertMsg = "找不到有效訂單"
                    Case Else
                        alertMsg = "No valid Delivery Order Available!"
                End Select

                ClientScript.RegisterStartupScript(Me.GetType(), "JSFUN", "alert('" & alertMsg & "');", True)

            End If


            load_ModalPopupExtender.Hide()
        Else
            Dim alertMsg As String = ""
            Select Case Session("gLang")
                Case "C"
                    alertMsg = "請輸入客戶訂單號碼"
                Case Else
                    alertMsg = "Please enter Customer Order No."
            End Select

            ClientScript.RegisterStartupScript(Me.GetType(), "JSFUN", "alert('" & alertMsg & "');", True)
        End If



    End Sub
End Class
