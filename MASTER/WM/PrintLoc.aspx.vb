Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient


Partial Class MASTER_WM_PrintLoc
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Const pageAbb As String = "exportLOCEXL"


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Not IsPostBack Then
            ViewState("WH_CODE") = ""

            ViewState("WH_CODE") = Server.UrlDecode(Request("wh_code"))


            BindGV()

        End If
        ar.hideForm(Me)


    End Sub

    Private Sub BindGV()

        Dim dt As New DataTable
        Dim sqlString As String = ""

        Dim w_code As String = ViewState("WH_CODE")

        sqlString = " SELECT wms_warehouse.WH_NAME,  wms_wh_fl.FL_NAME,  wms_wh_area.AR_NAME, " & _
                    "  wms_wh_rack.RK_NAME,  wms_wh_bin.BN_CODE,  wms_wh_area.AR_CODE, " & _
                    "  wms_warehouse.WH_CODE,  wms_wh_fl.FL_NUM,  wms_wh_rack.RK_CODE,wms_wh_bin.bn_csms_code " & _
                    " FROM wms_warehouse,  wms_wh_fl,  wms_wh_area,  wms_wh_rack,  wms_wh_bin " & _
                    " WHERE wms_warehouse.IMP_CODE = wms_wh_fl.IMP_CODE AND wms_warehouse.WH_CODE    = wms_wh_fl.WH_CODE AND wms_wh_fl.IMP_CODE       = wms_wh_area.IMP_CODE " & _
                    " AND wms_wh_fl.WH_CODE        = wms_wh_area.WH_CODE AND wms_wh_fl.FL_NUM         = wms_wh_area.FL_NUM AND wms_wh_area.IMP_CODE     = wms_wh_rack.IMP_CODE " & _
                    " AND wms_wh_area.WH_CODE      = wms_wh_rack.WH_CODE AND wms_wh_area.FL_NUM       = wms_wh_rack.FL_NUM AND wms_wh_area.AR_CODE      = wms_wh_rack.AR_CODE " & _
                    " AND wms_wh_rack.IMP_CODE     = wms_wh_bin.IMP_CODE AND wms_wh_rack.WH_CODE      = wms_wh_bin.WH_CODE AND wms_wh_rack.FL_NUM       = wms_wh_bin.FL_NUM " & _
                    " AND wms_wh_rack.AR_CODE      = wms_wh_bin.AR_CODE AND wms_wh_rack.RK_CODE      = wms_wh_bin.RK_CODE " & _
                    " AND wms_warehouse.wh_code='" & gU.dbEncode(w_code) & "'" & _
                    " Order by wms_warehouse.WH_CODE,  wms_wh_fl.FL_NUM,  wms_wh_area.AR_CODE, wms_wh_rack.RK_CODE,  wms_wh_bin.BN_CODE "
        dt = gDB.getDataTable(sqlString)

        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()

    End Sub


    Protected Sub btnSel_Click(sender As Object, e As System.EventArgs) Handles btnSel.Click
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked = True
            Next
        End If
    End Sub

    Protected Sub btnUnSel_Click(sender As Object, e As System.EventArgs) Handles btnUnSel.Click
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked = False
            Next
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("WH_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WH_CODE").ToString.Trim
                CType(e.Row.FindControl("FL_NUM"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "FL_NUM").ToString.Trim
                CType(e.Row.FindControl("AR_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "AR_CODE").ToString.Trim
                CType(e.Row.FindControl("RK_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RK_CODE").ToString.Trim

                CType(e.Row.FindControl("WH_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "WH_NAME").ToString.Trim
                CType(e.Row.FindControl("FL_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "FL_NAME").ToString.Trim
                CType(e.Row.FindControl("AR_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "AR_NAME").ToString.Trim
                CType(e.Row.FindControl("RK_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RK_NAME").ToString.Trim
                CType(e.Row.FindControl("BN_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "BN_CODE").ToString.Trim
                CType(e.Row.FindControl("BN_CSMS_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "BN_CSMS_CODE").ToString.Trim

        End Select
    End Sub

    Protected Sub btnPrint3_Click(sender As Object, e As System.EventArgs) Handles btnPrint3.Click
        ExportExcel("N")
    End Sub

    Protected Sub btnPrint2_Click(sender As Object, e As System.EventArgs) Handles btnPrint2.Click
        ExportExcel()
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As System.EventArgs) Handles btnPrint1.Click
        ExportExcel()
    End Sub


    Protected Sub ExportExcel(Optional havebarcode As String = "Y")
        Dim tempDT As New DataTable
        Dim sqlString As String
        Dim tickCount As Integer = 0
        gU.clearSessionTempData(pageAbb)

        sqlString = " Select '' as WH_NAME,'' as FL_NAME,'' as AR_NAME,'' as RK_NAME,'' as BN_CODE,'' as LCN_CODE, '' as BN_CSMS_CODE  where 1=3"
        tempDT = gDB.getDataTable(sqlString)

        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then

            Dim nRow As DataRow


            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked Then
                    nRow = tempDT.NewRow
                    nRow.Item("WH_NAME") = CType(GridView1.Rows(i).FindControl("WH_NAME"), Label).Text
                    nRow.Item("FL_NAME") = CType(GridView1.Rows(i).FindControl("FL_NAME"), Label).Text
                    nRow.Item("RK_NAME") = CType(GridView1.Rows(i).FindControl("RK_NAME"), Label).Text
                    nRow.Item("AR_NAME") = CType(GridView1.Rows(i).FindControl("AR_NAME"), Label).Text
                    nRow.Item("BN_CODE") = CType(GridView1.Rows(i).FindControl("BN_CODE"), Label).Text
                    nRow.Item("BN_CSMS_CODE") = CType(GridView1.Rows(i).FindControl("BN_CSMS_CODE"), Label).Text

                    nRow.Item("LCN_CODE") = CType(GridView1.Rows(i).FindControl("WH_CODE"), HiddenField).Value & CType(GridView1.Rows(i).FindControl("FL_NUM"), HiddenField).Value.ToString.PadLeft(2, "0") & _
                                           CType(GridView1.Rows(i).FindControl("AR_CODE"), HiddenField).Value.ToString.PadLeft(3, "0") & _
                                           CType(GridView1.Rows(i).FindControl("RK_CODE"), HiddenField).Value.ToString.PadLeft(4, "0") & _
                                           CType(GridView1.Rows(i).FindControl("BN_CODE"), Label).Text.ToString.PadLeft(3, "0")
                    tempDT.Rows.Add(nRow)
                    tickCount += 1
                    'WH_CODE.Value & "" & FL_NUM.Value.ToString.PadLeft(2, "0") & "" & AR_CODE.Value.PadLeft(3, "0") & "" & RK_CODE.Value.ToString.PadLeft(4, "0") & "" & nDT.Rows(i).Item(i1).ToString.PadLeft(3, "0")
                End If
            Next

            If tickCount > 0 Then
                tempDT.AcceptChanges()
                gU.setSessionTempData(pageAbb, "codeDT", tempDT)

                Dim rmtPost As New RemotePost

                rmtPost.Add("WH_CODE", ViewState("WH_CODE"))
                rmtPost.Add("havebarcode", havebarcode)
                'rmtPost.alertMsg = "New Item was cloned Successfully. Forwarding to New Item."
                rmtPost.Url = "EXP_LOCEXCL.aspx"
                rmtPost.Post()

            Else
                uiFun.displayMsg(Me, "", "Please select any bin for export!", Session("gLang"))

            End If
        End If
    End Sub

End Class
