Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_D_DELI_NOTE_DriverDeliNote
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        If Not IsPostBack Then


            If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
            do_code = Request.Form("do_code")
            storer_code = Request.Form("storer_code")

            sqlString = "select dm.storer_code, st.sto_name, " & _
                        "Convert(varchar, dm.do_date, 103) as do_date, " & _
                        "dm.cus_name, " & _
                        "(case when not dm.do_addr1 is null then dm.do_addr1 else '' end + " & _
                        "case when not dm.do_addr2 is null  then char(13) + char(10) + dm.do_addr2 else '' end + " & _
                        "case when not dm.do_addr3 is null  then char(13) + char(10) + dm.do_addr3 else '' end + " & _
                        "case when not dm.do_area_del is null  then char(13) + char(10) + dm.do_area_del else '' end + " & _
                        "case when not dm.do_region_del is null  then ' ' + dm.do_region_del else '' end + " & _
                        "case when not dm.do_country_del is null  then char(13) + char(10) + dm.do_country_del else '' end) as do_addr, " & _
                        "convert(decimal, ISNULL(dm.do_totl_pallets, 0)) as total_pallet, " & _
                        "convert(decimal, ISNULL(dm.do_totl_cartons, 0)) as total_carton, " & _
                        "dm.do_cus_cont, " & _
                        "dm.do_cus_cont_tel, " & _
                        "dm.do_rem, " & _
                        "dm.do_issued_by, " & _
                        "dm.do_driver " & _
                        "from wms_delv_order dm, wms_delv_order_d dd,wms_storer st " & _
                        "where " & _
                        "dm.imp_code = dd.imp_code " & _
                        "and dm.storer_code = dd.storer_code " & _
                        "and dm.do_code = dd.do_code " & _
                        "and dm.storer_code=st.storer_code and dm.imp_code=st.imp_code " & _
                        "and dm.imp_code = '" & imp_code & "' " & _
                        "and dm.do_code='" & gU.dbEncode(do_code) & "' and dm.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        "order by dd.dod_disp_seq "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then
                reportSource(nDataSource)
            Else
                Response.Write("No Driver Delivery Note Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\D_DELI_NOTE\ddn.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))

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
End Class
