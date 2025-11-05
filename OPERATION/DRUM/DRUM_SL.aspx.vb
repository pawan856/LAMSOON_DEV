Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient


Partial Class OPERATION_DRUM_DRUM_SL
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans

    Private DDFORMAT As String = "DD/MM/YYYY"
    Private exceptionEditList As List(Of String)
    Private moduleAction As String = ""
    Private dt As New DataTable


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If


        If Not IsPostBack Then
            ViewState("dt") = Nothing


            ViewState("DRUM_ID") = ""
            ViewState("DRUM_ID") = Server.UrlDecode(Request("DRUM_ID"))

            ViewState("STORER_CODE") = ""
            ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

            Call BindGV()

        End If
    End Sub

    Public Sub BindGV()
        Dim selectSQL As String = ""
        Dim pkCode As String = ""
        Dim storerCode As String = ""


        pkCode = ViewState("DRUM_ID")
        storerCode = ViewState("STORER_CODE")
        DRUM_ID.Text = pkCode

        selectSQL = " SELECT WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM.PACK_KEY, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S.SYS_LUD, WMS_ITEM_LOC_BAL.ILOC_LOC, " & _
                    " WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_KG, WMS_ITEM_LOC_BAL_S.ILBS_SEQ, WMS_ITEM_LOC_BAL_S.SYS_LUD, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO,case when WMS_ITEM_LOC_BAL_S.ILBS_KG is not null then 'KG' else NULL end as ILBS_UOM3, " & _
                    " WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL " & _
                    " FROM WMS_ITEM_LOC_BAL_S INNER JOIN " & _
                    " WMS_ITEM_LOC_BAL ON WMS_ITEM_LOC_BAL_S.ILOC_SEQ = WMS_ITEM_LOC_BAL.ILOC_SEQ INNER JOIN " & _
                    " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                    " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                    " Where WMS_ITEM_LOC_BAL_S.STORER_CODE='" & gU.dbEncode(storerCode) & "' and WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID='" & gU.dbEncode(pkCode) & "'" & _
		    " AND WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0 " & _
                    " order by  WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL"

        dt = gDB.getDataTable(selectSQL)

        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()

    End Sub
End Class
