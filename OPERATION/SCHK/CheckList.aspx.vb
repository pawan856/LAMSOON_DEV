Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OPERATION_SCHK_CheckList
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim ck_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim tempSQL As String = ""
        Dim tempSQL2 As String = ""

        Dim Paras(7) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("ck_code") = ""
        ViewState("ck_code") = Request("ck_code")

        ViewState("storer_code") = ""
        ViewState("storer_code") = Request("storer_code")

        ViewState("sheet_type") = ""
        ViewState("sheet_type") = Request("sheet_type")

        If Not IsPostBack Then

            If ViewState("sheet_type") = "R" Then
                tempSQL = " AND d.CKD_STATUS='RECHECK' "
                tempSQL2 = " AND WMS_STOCK_CHECK_D.CKD_STATUS='RECHECK' "

                ViewState("sheet_type") = "N"

            End If

            storer_code = ViewState("storer_code")
            ck_code = ViewState("ck_code")

            sqlString = " select d.CKD_SERIAL, i.ITM_GP_CODE, m.ck_type as type_code,m.CK_PERIOD, m.ck_level, m.ck_item_level, d.ckd_itm_code, d.ckd_rev_qty, d.ckd_org_qty, d.ckd_loc, d.ckd_batch_no, i.itm_name,i.itm_desc, i.itm_sku_no, c.COLC_ENG_VALUE as ck_type, d.CKD_CC_LIST_NO, " & _
                        " i.itm_uom, concat(WMS_WH_BIN.WH_CODE, ' ', WMS_WH_BIN.FL_NUM,' ' , WMS_WH_BIN.AR_CODE,' ',BN_CSMS_CODE) as loc2, WMS_WH_BIN.WH_CODE,WMS_WAREHOUSE.WH_NAME, d.ckd_rem," & _
                        " case when i.itm_type='CABLE' then convert(varchar,isnull(d.ckd_actual_qty2, d.CKD_REV_QTY2)) when i.itm_type <> 'CABLE' and i.ITM_SERIAL_NO_YN = 'Y' then d.CKD_SERIAL_NO else convert(varchar,isnull(d.ckd_actual_qty, d.CKD_REV_QTY)) end as ckd_actual_qty, " & _
                        " d.CKD_SERIAL_NO, i.ITM_SERIAL_NO_YN, " & _
                        " case when i.itm_type='CABLE' then isnull(d.ckd_book_qty2, d.CKD_org_QTY2) else isnull(d.ckd_book_qty, d.CKD_org_QTY) end as ckd_book_qty, " & _
                        " case when i.itm_type='CABLE' then d.CKD_var_QTY2 else d.ckd_var_qty end as ckd_var_qty " & _
                        " from wms_stock_check m inner join wms_stock_check_d d " & _
                        " on m.imp_code = d.imp_code and m.storer_code = d.storer_code and m.ck_code = d.ck_code " & _
                        " left outer join wms_item i " & _
                        " on d.CKD_ITM_CODE = i.itm_code and d.imp_code = i.imp_code and d.storer_code = i.storer_code and d.CKD_PACK_KEY = i.pack_key " & _
                        " left outer join wms_col_code c " & _
                        " on m.ck_type=c.COLC_CODE AND c.COLC_TABCOL='WMS_STOCK_CHECK.CK_TYPE' " & _
                        " Left Outer JOIN WMS_WH_BIN ON d.IMP_CODE = WMS_WH_BIN.IMP_CODE AND d.CKD_LOC = WMS_WH_BIN.LOC_KEY " & _
                        " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_WH_BIN.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND WMS_WH_BIN.WH_CODE = WMS_WAREHOUSE.WH_CODE " & _
                        " WHERE m.imp_code = '" & gU.dbEncode(imp_code) & "' AND m.storer_code='" & gU.dbEncode(storer_code) & "' and m.ck_code='" & gU.dbEncode(ck_code) & "'" & tempSQL & _
                        " order by  WMS_WH_BIN.WH_CODE, WMS_WH_BIN.FL_NUM,WMS_WH_BIN.AR_CODE,WMS_WH_BIN.BN_CSMS_CODE, i.itm_sku_no "
            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                Dim tempDT As DataTable
                Dim oBmp As Bitmap
                Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR") & "\SCHK\" & ck_code
                Dim CC_LIST_NO As String = ""

                If Not IO.Directory.Exists(tempPath) Then
                    IO.Directory.CreateDirectory(tempPath)
                End If

                sqlString = "select distinct CKD_CC_LIST_NO  from WMS_STOCK_CHECK_D Where WMS_STOCK_CHECK_D.imp_code = '" & gU.dbEncode(imp_code) & "' " & tempSQL2 & _
                                " AND WMS_STOCK_CHECK_D.ck_code='" & gU.dbEncode(ck_code) & "' and WMS_STOCK_CHECK_D.storer_code='" & gU.dbEncode(storer_code) & "' order by CKD_CC_LIST_NO "

                tempDT = gDB.getDataTable(sqlString)

                If tempDT.Rows.Count > 0 Then
                    For i = 0 To tempDT.Rows.Count - 1

                        CC_LIST_NO = ck_code & "-" & tempDT.Rows(i).Item("CKD_CC_LIST_NO").ToString.Trim

                        oBmp = Code128Rendering.MakeBarcodeImage(CC_LIST_NO, 2, True)
                        oBmp.Save(tempPath & "\" & "CK" & ck_code & "_" & tempDT.Rows(i).Item("CKD_CC_LIST_NO").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                    Next
                End If

                Dim ck_type As String = ""
                Dim ck_wh As String = ""
                Dim nsDT As DataTable
                Dim ns_no As Integer = 0
                Dim s_no As Integer = 0
                Dim sheet_type As String = ""
                Dim CK_PERIOD As String = ""


                Dim whStr As String = ""

                For i = 0 To nDataSource.Rows.Count - 1
                    whStr = gU.appendToList(whStr, nDataSource.Rows(i).Item("WH_NAME").ToString.Trim)
                Next

                If nDataSource.Rows(0).Item("type_code").ToString.Trim = "CCC" Then
                    ck_wh = "CBL, Cable Warehouse"
                Else
                    ck_wh = whStr
                End If

                sqlString = " SELECT count(case when itm_serial_no_yn = 'Y' then 1 end) as cs_itm_no, count(case when isnull(itm_serial_no_yn,'N') = 'N' or itm_serial_no_yn = '' then 1 end) as ns_itm_no " & _
                            " FROM WMS_STOCK_CHECK_D INNER JOIN " & _
                            " WMS_ITEM ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                            " WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_CHECK_D.CKD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                            " Where WMS_STOCK_CHECK_D.imp_code = '" & gU.dbEncode(imp_code) & "' " & tempSQL2 & _
                            " AND WMS_STOCK_CHECK_D.ck_code='" & gU.dbEncode(ck_code) & "' and WMS_STOCK_CHECK_D.storer_code='" & gU.dbEncode(storer_code) & "'"

                nsDT = gDB.getDataTable(sqlString)

                If nsDT.Rows.Count > 0 Then
                    s_no = nsDT.Rows(0).Item("cs_itm_no").ToString.Trim()
                    ns_no = nsDT.Rows(0).Item("ns_itm_no").ToString.Trim()
                End If

                'ck_type = nDataSource.Rows(0).Item("ck_type").ToString.Trim

                'Select Case nDataSource.Rows(0).Item("ck_level").ToString.Trim
                '    Case "WH"
                '        ck_level = "Warehouse"
                '    Case "FL"
                '        ck_level = "Floor"
                '    Case "AR"
                '        ck_level = "Area"
                'End Select

                'Select Case nDataSource.Rows(0).Item("ck_item_level").ToString.Trim
                '    Case "AL"
                '        ck_item_level = "ALL"
                '    Case "IT"
                '        ck_item_level = "Item"
                'End Select
                CK_PERIOD = nDataSource.Rows(0).Item("CK_PERIOD").ToString.Trim
                sheet_type = ViewState("sheet_type")

                Paras(0) = New ReportParameter("CK_TYPE", ck_type)
                Paras(1) = New ReportParameter("CK_CODE", ck_code)
                Paras(2) = New ReportParameter("BarcodePath", "file:///" & tempPath & "\")
                Paras(3) = New ReportParameter("CK_WH", ck_wh)
                Paras(4) = New ReportParameter("NS_ITEM_NO", ns_no)
                Paras(5) = New ReportParameter("S_ITEM_NO", s_no)
                Paras(6) = New ReportParameter("SHEET_TYPE", sheet_type)
                Paras(7) = New ReportParameter("CK_PERIOD", CK_PERIOD)

                reportSource(nDataSource, Paras)
            Else

                Response.Write("No Check List record has been Found.")
                Response.End()
            End If


        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
            End If

            Try

                ReportViewer1.LocalReport.Refresh()

                Dim formatName As String = "word"

                For Each extension As RenderingExtension In ReportViewer1.LocalReport.ListRenderingExtensions

                    If extension.Name.ToLower = formatName Then

                        Dim m_isVisible As System.Reflection.FieldInfo = extension.GetType.GetField("m_isVisible", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)

                        m_isVisible.SetValue(extension, False)

                        Exit For

                    End If

                Next



            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
