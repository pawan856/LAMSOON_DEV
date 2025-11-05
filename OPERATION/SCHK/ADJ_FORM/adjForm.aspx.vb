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

        Dim Paras(2) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("ck_code") = ""
        ViewState("ck_code") = Request("ck_code")

        ViewState("storer_code") = ""
        ViewState("storer_code") = Request("storer_code")

        If Not IsPostBack Then

            storer_code = ViewState("storer_code")
            ck_code = ViewState("ck_code")

            sqlString = "SELECT WMS_STOCK_CHECK_D.ckd_seq, WMS_STOCK_CHECK_D.CKD_ORG_QTY, WMS_STOCK_CHECK_D.CKD_REV_QTY, WMS_STOCK_CHECK_D.CKD_VAR_QTY, " & _
                        " WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY, WMS_STOCK_CHECK_D.CKD_ORG_QTY2, WMS_STOCK_CHECK_D.CKD_REV_QTY2,  WMS_WH_BIN.BN_CSMS_CODE as CKD_LOC, " & _
                        " WMS_STOCK_CHECK_D.CKD_VAR_QTY2, WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY2, WMS_STOCK_CHECK_D.CKD_BOOK_QTY,  " & _
                        " WMS_STOCK_CHECK_D.CKD_BOOK_QTY2, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_GP_CODE,WMS_STOCK_CHECK_D.CKD_REM " & _
                        " FROM WMS_ITEM INNER JOIN " & _
                        " WMS_STOCK_CHECK_D ON WMS_ITEM.IMP_CODE = WMS_STOCK_CHECK_D.IMP_CODE AND  " & _
                        " WMS_ITEM.STORER_CODE = WMS_STOCK_CHECK_D.STORER_CODE AND WMS_ITEM.ITM_CODE = WMS_STOCK_CHECK_D.CKD_ITM_CODE AND  " & _
                        " WMS_ITEM.PACK_KEY = WMS_STOCK_CHECK_D.CKD_PACK_KEY INNER JOIN " & _
                        " WMS_STOCK_CHECK ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_STOCK_CHECK.IMP_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.STORER_CODE = WMS_STOCK_CHECK.STORER_CODE AND WMS_STOCK_CHECK_D.CK_CODE = WMS_STOCK_CHECK.CK_CODE " & _
                        " LEFT OUTER JOIN WMS_WH_BIN ON WMS_STOCK_CHECK_D.CKD_LOC = WMS_WH_BIN.LOC_KEY " & _
                        " WHERE WMS_STOCK_CHECK_D.storer_code='" & gU.dbEncode(storer_code) & "' and WMS_STOCK_CHECK_D.ck_code='" & gU.dbEncode(ck_code) & "'" & _
                        " AND ((WMS_STOCK_CHECK_D.CKD_VAR_QTY is not null and WMS_STOCK_CHECK_D.CKD_VAR_QTY <> 0) or (WMS_STOCK_CHECK_D.CKD_VAR_QTY2 is not null and WMS_STOCK_CHECK_D.CKD_VAR_QTY2 <> 0))" & _
                        " order by WMS_STOCK_CHECK_D.ckd_seq"

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then
                'Dim ck_type As String = ""
                'Dim ck_level As String = ""
                'Dim ck_item_level As String = ""

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

                'Paras(0) = New ReportParameter("CK_TYPE", ck_type)
                'Paras(1) = New ReportParameter("CK_LEVEL", ck_level)
                'Paras(2) = New ReportParameter("CK_ITEM_LEVEL", ck_item_level)

                reportSource(nDataSource, Paras)
            Else

                Response.Write("No adjustment record has been Found.")
                Response.End()
            End If


        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

            'If Not IsNothing(paraarray) Then
            '    ReportViewer1.LocalReport.SetParameters(paraarray)
            'End If

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
