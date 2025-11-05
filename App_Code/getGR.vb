Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization
Imports System.Data
Imports System.Data.SqlClient

Public Class IGRResponse
    <XmlElement(ElementName:="gr_code")> _
   Public gr_code As String

    <XmlElement(ElementName:="storer_code")> _
    Public storer_code As String

    <XmlElement(ElementName:="gr_status")> _
    Public gr_status As String

    <XmlElement(ElementName:="gr_date")> _
    Public gr_date As String

    <XmlElement(ElementName:="gr_doc_no")> _
    Public gr_doc_no As String

    <XmlElement(ElementName:="gr_tot_pallet")> _
    Public gr_tot_pallet As String

    <XmlElement(ElementName:="gr_track_no")> _
    Public gr_track_no As String

    <XmlElement(ElementName:="gr_ref_no")> _
    Public gr_ref_no As String

    <XmlElement(ElementName:="grd_seq")> _
    Public grd_seq As String

    <XmlElement(ElementName:="grd_disp_seq")> _
    Public grd_disp_seq As String

    <XmlElement(ElementName:="grd_batch_no")> _
    Public grd_batch_no As String

    <XmlElement(ElementName:="grd_itm_code")> _
    Public grd_itm_code As String

    <XmlElement(ElementName:="grd_pack_key")> _
    Public grd_pack_key As String

    <XmlElement(ElementName:="grd_itm_name")> _
    Public grd_itm_name As String

    <XmlElement(ElementName:="grd_batch_cd")> _
    Public grd_batch_cd As String

    <XmlElement(ElementName:="grd_po_qty")> _
    Public grd_po_qty As String

    <XmlElement(ElementName:="grd_rcv_qty")> _
    Public grd_rcv_qty As String

    <XmlElement(ElementName:="grd_os_qty")> _
    Public grd_os_qty As String

    <XmlElement(ElementName:="grd_width")> _
    Public grd_width As String

    <XmlElement(ElementName:="grd_length")> _
    Public grd_length As String

    <XmlElement(ElementName:="grd_height")> _
    Public grd_height As String

    <XmlElement(ElementName:="grd_kg")> _
    Public grd_kg As String

    <XmlElement(ElementName:="grd_cbm")> _
    Public grd_cbm As String

    <XmlElement(ElementName:="do_job_id")> _
    Public do_job_id As String
End Class

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class getGR
    Inherits System.Web.Services.WebService

    Private gU As New GeneralUtils
    Private gDB As New GlobalDBFunc

    <WebMethod()> _
  Public Function getGRList(ByVal gr_code As String) As IGRResponse()
        Dim nSQL As String
        Dim datatbl As New DataTable

        nSQL = "select gr.imp_code, gr.storer_code, gr.gr_code, gr.gr_status, gr.gr_date, gr.do_job_id, " & _
                "gr.gr_doc_no, gr.gr_tot_pallet, gr.gr_track_no, gr.gr_ref_no, " & _
                "d.grd_seq,d.grd_disp_seq,d.grd_batch_no,d.grd_itm_code," & _
                "d.grd_pack_key,d.grd_itm_name,d.grd_batch_cd,d.grd_po_qty," & _
                "d.grd_rcv_qty,d.grd_os_qty,d.grd_width,d.grd_length,d.grd_height," & _
                "d.grd_kg,d.grd_cbm " & _
                   "from wms_goodsrcv gr, wms_goodsrcv_d d " & _
                   "where " & _
                   "gr.imp_code = d.imp_code " & _
                   "and gr.storer_code = d.storer_code " & _
                   "and gr.gr_code = d.gr_code " & _
                   "and gr.gr_code = '" & gr_code & "' " & _
                   "and gr.imp_code = 'WMS'"

        nSQL = nSQL & " order by wms_goodsrcv_d_s.grs_disp_seq"

        Try
            datatbl = getDataTable(nSQL)

            Dim CodeResponse(datatbl.Rows.Count - 1) As IGRResponse

            For i As Integer = 0 To datatbl.Rows.Count - 1
                CodeResponse(i) = New IGRResponse
                CodeResponse(i).storer_code = datatbl.Rows(i).Item("storer_code").ToString
                CodeResponse(i).do_job_id = datatbl.Rows(i).Item("do_job_id").ToString
                CodeResponse(i).gr_code = datatbl.Rows(i).Item("gr_code").ToString
                CodeResponse(i).gr_status = datatbl.Rows(i).Item("gr_status").ToString
                CodeResponse(i).gr_date = datatbl.Rows(i).Item("gr_date").ToString
                CodeResponse(i).gr_doc_no = datatbl.Rows(i).Item("gr_doc_no").ToString
                CodeResponse(i).gr_tot_pallet = datatbl.Rows(i).Item("gr_tot_pallet").ToString
                CodeResponse(i).gr_track_no = datatbl.Rows(i).Item("gr_track_no").ToString
                CodeResponse(i).gr_ref_no = datatbl.Rows(i).Item("gr_ref_no").ToString
                CodeResponse(i).grd_seq = datatbl.Rows(i).Item("grd_seq").ToString
                CodeResponse(i).grd_disp_seq = datatbl.Rows(i).Item("grd_disp_seq").ToString
                CodeResponse(i).grd_batch_no = datatbl.Rows(i).Item("grd_batch_no").ToString
                CodeResponse(i).grd_itm_code = datatbl.Rows(i).Item("grd_itm_code").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_pack_key").ToString

                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_itm_name").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_batch_cd").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_po_qty").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_rcv_qty").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_os_qty").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_width").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_length").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_height").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_kg").ToString
                CodeResponse(i).grd_pack_key = datatbl.Rows(i).Item("grd_cbm").ToString
            Next

            Return CodeResponse

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function getConnection() As SqlConnection
        Dim l_conn As SqlConnection

        l_conn = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        l_conn.Open()

        Return l_conn

    End Function

    Private Function getDataTable(ByVal as_sql As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing, Optional ByRef TableSchema As DataTable = Nothing) As DataTable
        Dim lds_1 As New DataSet
        Dim lda_1 As SqlDataAdapter
        Dim l_conn As SqlConnection
        Dim l_SqlCmd As SqlCommand

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If

        l_SqlCmd = New SqlCommand(as_sql, l_conn)
        If Not pTransaction Is Nothing Then
            l_SqlCmd.Transaction = pTransaction
        End If

        lda_1 = New SqlDataAdapter(l_SqlCmd)

        If Not TableSchema Is Nothing Then lda_1.FillSchema(TableSchema, SchemaType.Source)

        lda_1.Fill(lds_1)

        If pConn Is Nothing Then
            l_conn.Close()
            l_conn.Dispose()
        End If

        Return lds_1.Tables(0)

    End Function
End Class
