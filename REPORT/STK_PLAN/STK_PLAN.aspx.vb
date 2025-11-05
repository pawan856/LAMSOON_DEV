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
        ar = New AccessRightUtils("RPT_STK_PLAN", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim Paras(2) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Dim tempSQL As String = ""
            storer_code = ViewState("storer_code")

            'If Not String.IsNullOrWhiteSpace(Session("SEARCH_SESSION_PAGE_V_LOCATION_WH_CODE")) Then

            'End If

            If Not String.IsNullOrWhiteSpace(Session("GENERIC_SESSION_SRCH_WHERE_SQL")) Then
                tempSQL = " AND " & Session("GENERIC_SESSION_SRCH_WHERE_SQL")
            End If

            sqlString = " select a.WH_CODE, a.FL_NUM, a.AR_CODE, WMS_COL_CODE.COLC_ENG_VALUE as item_price_class, a.itm_count, " & _
                        " a.CK_START_DATE, a.CK_END_DATE,convert(varchar,a.CK_START_DATE,103) as ck_start_date_D, convert(varchar,a.CK_END_DATE,103) as ck_end_date_D, gp_code =  " & _
                        " STUFF((SELECT ', ' + itm_gp_code " & _
                        " FROM WMS_STOCK_CHECK_D INNER JOIN " & _
                        " WMS_ITEM ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_CHECK_D.CKD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                        " WMS_STOCK_CHECK ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_STOCK_CHECK.IMP_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.STORER_CODE = WMS_STOCK_CHECK.STORER_CODE AND " & _
                        " WMS_STOCK_CHECK_D.CK_CODE = WMS_STOCK_CHECK.CK_CODE INNER JOIN " & _
                        " V_LOCATION ON WMS_STOCK_CHECK_D.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_STOCK_CHECK_D.CKD_LOC = V_LOCATION.LOC " & _
                        " WHERE V_LOCATION.WH_CODE = a.WH_CODE and V_LOCATION.FL_NUM = a.fl_num and V_LOCATION.AR_CODE =a.ar_code and WMS_ITEM.ITEM_PRICE_CLASS = a.item_price_class and " & _
                        " WMS_STOCK_CHECK.CK_START_DATE =a.ck_start_date and  WMS_STOCK_CHECK.CK_END_DATE = a.ck_end_date " & _
                        " group by  V_LOCATION.WH_CODE, V_LOCATION.FL_NUM, V_LOCATION.AR_CODE, WMS_ITEM.ITEM_PRICE_CLASS, " & _
                        " WMS_STOCK_CHECK.CK_START_DATE, WMS_STOCK_CHECK.CK_END_DATE, wms_item.itm_gp_code " & _
                        " FOR XML PATH(''), TYPE).value('.', 'varchar(max)'),1,1,''),  " & _
                        " checker = STUFF((SELECT ', ' + CKD_CHECKER  " & _
                        " FROM WMS_STOCK_CHECK_D INNER JOIN  " & _
                        " WMS_ITEM ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_CHECK_D.CKD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN  " & _
                        " WMS_STOCK_CHECK ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_STOCK_CHECK.IMP_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.STORER_CODE = WMS_STOCK_CHECK.STORER_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.CK_CODE = WMS_STOCK_CHECK.CK_CODE INNER JOIN  " & _
                        " V_LOCATION ON WMS_STOCK_CHECK_D.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_STOCK_CHECK_D.CKD_LOC = V_LOCATION.LOC " & _
                        " WHERE V_LOCATION.WH_CODE = a.WH_CODE and V_LOCATION.FL_NUM = a.fl_num and V_LOCATION.AR_CODE =a.ar_code and WMS_ITEM.ITEM_PRICE_CLASS = a.item_price_class and " & _
                        " WMS_STOCK_CHECK.CK_START_DATE =a.ck_start_date and  WMS_STOCK_CHECK.CK_END_DATE = a.ck_end_date " & _
                        " group by  V_LOCATION.WH_CODE, V_LOCATION.FL_NUM, V_LOCATION.AR_CODE, WMS_ITEM.ITEM_PRICE_CLASS, " & _
                        " WMS_STOCK_CHECK.CK_START_DATE, WMS_STOCK_CHECK.CK_END_DATE, WMS_STOCK_CHECK_D.CKD_CHECKER " & _
                        " FOR XML PATH(''), TYPE).value('.', 'varchar(max)'),1,1,'') " & _
                        " from (SELECT  V_LOCATION.WH_CODE, V_LOCATION.FL_NUM, V_LOCATION.AR_CODE, WMS_ITEM.ITEM_PRICE_CLASS, count(distinct WMS_ITEM.ITM_SKU_NO) as itm_count, " & _
                        " WMS_STOCK_CHECK.CK_START_DATE, WMS_STOCK_CHECK.CK_END_DATE " & _
                        " FROM WMS_STOCK_CHECK_D INNER JOIN " & _
                        " WMS_ITEM ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_CHECK_D.CKD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                        " WMS_STOCK_CHECK ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_STOCK_CHECK.IMP_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.STORER_CODE = WMS_STOCK_CHECK.STORER_CODE AND  " & _
                        " WMS_STOCK_CHECK_D.CK_CODE = WMS_STOCK_CHECK.CK_CODE INNER JOIN " & _
                        " V_LOCATION ON WMS_STOCK_CHECK_D.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_STOCK_CHECK_D.CKD_LOC = V_LOCATION.LOC " & _
                        " where 1=1 " & _
                        tempSQL & _
                        " group by  V_LOCATION.WH_CODE, V_LOCATION.FL_NUM, V_LOCATION.AR_CODE, WMS_ITEM.ITEM_PRICE_CLASS,  " & _
                        " WMS_STOCK_CHECK.CK_START_DATE, WMS_STOCK_CHECK.CK_END_DATE " & _
                        ") a " & _
                        " left outer join wms_col_code on a.ITEM_PRICE_CLASS = WMS_COL_CODE.COLC_CODE and WMS_COL_CODE.COLC_TABCOL='WMS_ITEM.ITEM_PRICE_CLASS' " & _
                        " ORDER BY a.WH_CODE, a.FL_NUM, a.AR_CODE, item_price_class "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then
                reportSource(nDataSource, Paras)
            Else
                Response.Write("&nbsp;&nbsp;<input type=""button"" value=""Back"" class=""all_button"" onclick=""Javascript:window.location='../../cms_search.aspx?menu_code=RPT_STK_PLAN'"" />")
                Response.Write("<br /><br />&nbsp;&nbsp;No Cycle Count Plan has been Found.")
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
