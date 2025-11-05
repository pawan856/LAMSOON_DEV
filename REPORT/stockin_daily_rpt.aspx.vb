Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization

Imports Microsoft.Reporting.WebForms

Partial Class REPORT_stockin_daily_rpt
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gFunc As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils("RPT_SID", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        'ar.hideForm(Me)

        If Not IsPostBack Then
            BindGV(IsPostBack)
        End If
    End Sub

    Private Sub BindGV(ByVal isPostBack As Boolean)
        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim gr_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim gr_doc_type As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        Dim fromdate As DateTime
        Dim todate As DateTime
        Dim vfromdate As String = ""
        Dim vtodate As String = ""
        Dim vStorer As String = ""
        Dim vRef As String = ""

        Dim vTrackNo As String = ""

        Dim s As String()
        Dim st As String()

        Dim vWhereSQL As String = ""
        Dim vGroupBySQL As String = ""
        Dim vOrderSQL As String = ""

        If isPostBack Then
            vfromdate = ViewState("vfromdate")
            fromdate = ViewState("fromdate")
            todate = ViewState("todate")
            vtodate = ViewState("vtodate")
            vTrackNo = ViewState("vTrackNo")
            vStorer = ViewState("vStorer")
            sqlString = ViewState("sqlString")
            vWhereSQL = " Where " & ViewState("vWhereSQL")
            vGroupBySQL = ViewState("vGroupBySQL")
            vOrderSQL = ViewState("vOrderSQL")
        Else
            If Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_DATE") IsNot Nothing Then
                fromdate = New DateTime()
                Dim provider As CultureInfo = CultureInfo.InvariantCulture

                If Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_DATE") <> "" Then
                    fromdate = DateTime.ParseExact(Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_DATE"), "dd/MM/yyyy", provider)

                    'todate = fromdate.Date.AddMonths(1)
                    todate = fromdate.Date.AddDays(Date.DaysInMonth(fromdate.Year, fromdate.Month) - 1)

                    s = fromdate.Date.GetDateTimeFormats()
                    st = todate.Date.GetDateTimeFormats()

                    'vfromdate = s(10).ToString()
                    'vtodate = st(10).ToString()
                    vfromdate = fromdate.ToString("dd/MM/yyyy")
                    vtodate = todate.ToString("dd/MM/yyyy")
                Else
                    fromdate = Nothing
                    todate = Nothing
                End If

                ViewState("fromdate") = fromdate
                ViewState("todate") = todate
                ViewState("vfromdate") = vfromdate
                ViewState("vtodate") = vtodate

                'fromdate = DateTime.ParseExact(fromdate, "yyyy/MM/dd", Nothing)
            Else
                fromdate = ViewState("fromdate")
                todate = ViewState("todate")
                vfromdate = ViewState("vfromdate")
                vtodate = ViewState("vtodate")
            End If

            If Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_TRACK_NO") IsNot Nothing Then

                If Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_TRACK_NO") <> "" Then
                    vTrackNo = Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_GR_TRACK_NO")
                Else
                    vTrackNo = ""
                End If

                ViewState("vTrackNo") = vTrackNo
            Else
                vTrackNo = ViewState("vTrackNo")
            End If

            If Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_STORER_CODE") IsNot Nothing Then
                vStorer = Session("SEARCH_SESSION_PAGE_WMS_GOODSRCV_STORER_CODE")
                ViewState("vStorer") = vStorer
            Else
                vStorer = ViewState("vStorer")
            End If
 
            If Session("GENERIC_SESSION_SRCH_SQL") IsNot Nothing Then
                sqlString = Session("GENERIC_SESSION_SRCH_SQL")
                ViewState("sqlString") = sqlString
            Else
                sqlString = ViewState("sqlString")
            End If


            If Session("GENERIC_SESSION_SRCH_WHERE_SQL") IsNot Nothing Then
                vWhereSQL = gU.decodeNullOrEmpty(Session("GENERIC_SESSION_SRCH_WHERE_SQL"), " WHERE 1 = 1 ")
                ViewState("vWhereSQL") = vWhereSQL
            Else
                vWhereSQL = ViewState("vWhereSQL")
            End If


            If Session("GENERIC_SESSION_SRCH_ADD_SQL") IsNot Nothing Then
                vGroupBySQL = Session("GENERIC_SESSION_SRCH_ADD_SQL")
                ViewState("vGroupBySQL") = vGroupBySQL
            Else
                vGroupBySQL = ViewState("vGroupBySQL")
            End If

            If Session("GENERIC_SESSION_SRCH_ORDERBY_SQL") IsNot Nothing Then
                vOrderSQL = Session("GENERIC_SESSION_SRCH_ORDERBY_SQL")
                ViewState("vOrderSQL") = vOrderSQL
            Else
                vOrderSQL = ViewState("vOrderSQL")
            End If

        End If

        If vfromdate = "" And vTrackNo = "" Then
            Dim strScript As String = "alert(""Date or Tracking No. Cannot Be Empty."");window.open('','_self');window.close();"

            If Not Me.ClientScript Is Nothing Then

                If Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType(), "") Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType(), "", strScript, True)
                End If
            End If

            Exit Sub
        End If

        Dim grAppSQL As String = ""

        If isPostBack = False Then

            If vWhereSQL <> "" Then vWhereSQL = " Where " & vWhereSQL

            Dim grString As String = "select distinct WMS_GOODSRCV.GR_CODE, WMS_GOODSRCV.GR_REF_NO " & _
                                "FROM WMS_GOODSRCV INNER JOIN WMS_GOODSRCV_D ON " & _
                                    "WMS_GOODSRCV.GR_CODE = WMS_GOODSRCV_D.GR_CODE AND " & _
                                    "WMS_GOODSRCV.IMP_CODE = WMS_GOODSRCV_D.IMP_CODE AND " & _
                                    "WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE " & _
                                    " " & vWhereSQL



            Dim listDT As New DataTable

            listDT = gDB.getDataTable(grString)

            DataList1.DataSource = listDT
            DataList1.DataBind()
        Else
            For m As Integer = 0 To DataList1.Items.Count - 1
                If DirectCast(DataList1.Items(m).FindControl("gr_code_cb"), CheckBox).Checked Then
                    If grAppSQL <> "" Then grAppSQL = grAppSQL & ","

                    grAppSQL = grAppSQL & DirectCast(DataList1.Items(m).FindControl("gr_code_cb"), CheckBox).Text
                End If
            Next
        End If

        If vWhereSQL <> "" Then sqlString = sqlString & " " & vWhereSQL
        If grAppSQL <> "" Then sqlString = sqlString & " AND WMS_GOODSRCV.GR_CODE IN ('" & grAppSQL.Replace(",", "','") & "')"
        If vGroupBySQL <> "" Then sqlString = sqlString & " " & vGroupBySQL
        If vOrderSQL <> "" Then sqlString = sqlString & " " & vOrderSQL

        nDataSource = gDB.getDataTable(sqlString)

        If nDataSource.Rows.Count > 0 Then

            If nDataSource.Rows.Count = 1 Then
                vStorer = nDataSource.Rows(0).Item("STORER_CODE")
            Else
                If vStorer = "" Then
                    Dim hasDiff As Boolean = False
                    Dim firstSC As String = ""

                    For i As Integer = 0 To nDataSource.Rows.Count - 1
                        If firstSC = "" Then firstSC = nDataSource.Rows(i).Item("STORER_CODE").ToString

                        If nDataSource.Rows(i).Item("STORER_CODE").ToString <> firstSC Then
                            hasDiff = True
                            Exit For
                        End If
                    Next

                    If hasDiff = False Then
                        vStorer = firstSC
                    End If
                End If
            End If

            If vStorer = "" Then
                vStorer = "All"
                vRef = "ALL"
            Else
                vRef = vStorer

                Dim STORERSQL = "SELECT STO_NAME FROM WMS_STORER WHERE STORER_CODE = '" & vStorer & "'"

                vStorer = gFunc.getValueFromSQL(STORERSQL)
            End If

            If vWhereSQL <> "" Then vWhereSQL = vWhereSQL & " AND WMS_GOODSRCV_D.GRD_NO_OF_CARTON > 0 " Else vWhereSQL = "WHERE WMS_GOODSRCV_D.GRD_NO_OF_CARTON > 0 "

            Dim carString As String = "SELECT " & _
                                        "WMS_GOODSRCV.STORER_CODE, " & _
                                        "WMS_GOODSRCV.GR_DATE, " & _
                                        "WMS_GOODSRCV.GR_REF_NO, " & _
                                        "WMS_GOODSRCV_D.GRD_CARTON_NO, " & _
                                        "WMS_GOODSRCV_D.GRD_NO_OF_CARTON as ITM_CARTON, " & _
                                        "WMS_GOODSRCV_D.GRD_LENGTH, WMS_GOODSRCV_D.GRD_WIDTH, " & _
                                        "WMS_GOODSRCV_D.GRD_HEIGHT, " & _
                                        "(CASE WHEN ISNULL(WMS_GOODSRCV_D.GRD_NO_OF_CARTON,0) > 0 THEN WMS_GOODSRCV_D.GRD_CBM ELSE 0 END) AS GRD_CBM, " & _
                                        "(CASE WHEN ISNULL(WMS_GOODSRCV_D.GRD_NO_OF_CARTON,0) > 0 THEN WMS_GOODSRCV_D.GRD_KG ELSE 0 END) AS GRD_KG, " & _
                                        "WMS_GOODSRCV_D.GRD_KG * ISNULL(WMS_GOODSRCV_D.GRD_NO_OF_CARTON,0) AS SUB_KG " & _
                                        "FROM WMS_GOODSRCV INNER JOIN WMS_GOODSRCV_D ON " & _
                                        "WMS_GOODSRCV.GR_CODE = WMS_GOODSRCV_D.GR_CODE AND " & _
                                        "WMS_GOODSRCV.IMP_CODE = WMS_GOODSRCV_D.IMP_CODE AND " & _
                                        "WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE " & _
                                        vWhereSQL

            If grAppSQL <> "" Then carString = carString & " AND WMS_GOODSRCV.GR_CODE IN ('" & grAppSQL.Replace(",", "','") & "') "

            carString = carString & " ORDER BY WMS_GOODSRCV.GR_DATE, WMS_GOODSRCV.GR_REF_NO, WMS_GOODSRCV_D.GRD_CARTON_NO "

            Dim carDataTable As New DataTable

            carDataTable = gDB.getDataTable(carString)

            For n As Integer = 0 To nDataSource.Rows.Count - 1
                Dim objSum As Object
                Dim rCBM As Object
                Dim rKG As Object

                objSum = carDataTable.Compute("Sum(ITM_CARTON)", "GR_REF_NO = '" & nDataSource.Rows(n).Item("GR_REF_NO") & "'")
                rCBM = carDataTable.Compute("Sum(GRD_CBM)", "GR_REF_NO = '" & nDataSource.Rows(n).Item("GR_REF_NO") & "'")
                rKG = carDataTable.Compute("Sum(SUB_KG)", "GR_REF_NO = '" & nDataSource.Rows(n).Item("GR_REF_NO") & "'")


                'Dim rows As DataRow() = carDataTable.Select("GR_REF_NO = '" & nDataSource.Rows(n).Item("GR_REF_NO") & "'", "")

                'Dim rCBM As Decimal = 0
                'Dim rKG As Decimal = 0

                'For g As Integer = 0 To rows.Count - 1
                '    Dim cbm As Double = gU.decodeEmptyCdbl(rows(g).Item("grd_length").ToString, 0) * _
                '                           gU.decodeEmptyCdbl(rows(g).Item("grd_width").ToString, 0) * _
                '                           gU.decodeEmptyCdbl(rows(g).Item("grd_height").ToString, 0)

                '    If cbm <> 0 Then cbm = cbm / 1000000

                '    'rCBM = rCBM + cbm * gU.decodeEmptyCInt(rows(g).Item("ITM_CARTON").ToString.Trim, 0)
                '    'rKG = rKG + gU.decodeEmptyCdbl(rows(g).Item("GRD_KG").ToString.Trim, 0) * gU.decodeEmptyCInt(rows(g).Item("ITM_CARTON").ToString.Trim, 0)

                '    rCBM = rCBM + cbm
                '    rKG = rKG + gU.decodeEmptyCdbl(rows(g).Item("GRD_KG").ToString.Trim, 0)
                'Next

                nDataSource.Rows(n).Item("TOTAL_CARTON") = objSum
                nDataSource.Rows(n).Item("GRD_CBM") = rCBM
                nDataSource.Rows(n).Item("GRD_KG") = rKG
            Next

            nDataSource.AcceptChanges()

            Dim paras(5) As ReportParameter

            vRef = vRef.Trim & Now.Date.ToString("yyyyMMdd")
            Dim vETA As String

            If vfromdate <> "" Then
                vETA = Format(fromdate, "dd/MM/yyyy")
            Else
                vETA = ""
            End If

            paras(0) = New ReportParameter("FROM_DATE", vfromdate)
            paras(1) = New ReportParameter("STORER", vStorer)
            paras(2) = New ReportParameter("REF", vRef)
            'paras(3) = New ReportParameter("CARTON", totalCarton)
            paras(3) = New ReportParameter("CARTON", "")
            paras(4) = New ReportParameter("ETA", vETA)
            paras(5) = New ReportParameter("TO_DATE", vtodate)

            reportSource(nDataSource, paras)

        Else
            Dim strScript As String = "alert(""No Stock In Record within this Date Range: " & gU.decodeNullOrEmpty(vfromdate, "All Date") & "."");window.open('','_self');window.close();"

            If Not Me.ClientScript Is Nothing Then

                If Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType(), "") Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType(), "", strScript, True)
                End If
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\stockin_daily_rpt.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_SI_DAILY", sourcetbl))

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
            End If

            Try
                ReportViewer1.LocalReport.Refresh()

            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        BindGV(True)
    End Sub
End Class
