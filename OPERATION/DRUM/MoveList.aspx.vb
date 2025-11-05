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
    Private DDFORMAT As String

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        ar.hideForm(Me)
        DDFORMAT = gU.getConfig("DDFORMATNO")
        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim DRUM_ID As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim Paras(2) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("DRUM_ID") = ""
        ViewState("DRUM_ID") = Request("DRUM_ID")

        ViewState("storer_code") = ""
        ViewState("storer_code") = Request("storer_code")

        If Not IsPostBack Then

            storer_code = ViewState("storer_code")
            DRUM_ID = ViewState("DRUM_ID")

            sqlString = " SELECT WMS_DRUM_DTL.IMP_CODE, WMS_DRUM_DTL.STORER_CODE, WMS_DRUM_DTL.DRUM_ID, WMS_DRUM_DTL.DRUD_SEQ, DRUD_DATE as full_DRUD_DATE, convert(varchar,DRUD_DATE," & DDFORMAT & ") as DRUD_DATE,c1.colc_eng_value as DRUD_MOVEMENT, WMS_COL_CODE.colc_eng_value as DRUD_DOC_TYPE, DRUD_DOC_NO, DRUD_BY, DRUD_LOC, DRUD_MV_TYPE, " & _
                    " DRUD_REMARKS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " & _
                    " V_LOCATION.WH_CODE as DRUD_WH " & _
                    " FROM WMS_DRUM_DTL " & _
                    " LEFT OUTER JOIN V_LOCATION ON WMS_DRUM_DTL.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_DRUM_DTL.DRUD_LOC = V_LOCATION.LOC " & _
                    " left JOIN WMS_COL_CODE AS c1 ON WMS_DRUM_DTL.DRUD_MOVEMENT = c1.COLC_CODE and c1.COLC_TABCOL='WMS_DRUM_DTL.DRUD_MOVEMENT'" & _
                    " left outer JOIN WMS_COL_CODE ON WMS_DRUM_DTL.DRUD_DOC_TYPE = WMS_COL_CODE.COLC_CODE and WMS_COL_CODE.COLC_TABCOL='WMS_DRUM_DTL.DRUD_DOC_TYPE' " & _
                    " where WMS_DRUM_DTL.DRUM_ID = '" & gU.dbEncode(DRUM_ID) & "' " & _
                    " and WMS_DRUM_DTL.storer_code = '" & gU.dbEncode(storer_code) & "' " & _
                    " and WMS_DRUM_DTL.imp_code = '" & Session("IMP_CODE") & "'" & _
                    " order by CONVERT(int, WMS_DRUM_DTL.DRUD_SEQ)"

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

                Response.Write("No movement record has been Found.")
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
