Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class INBOUND_GR_GR_LABELS_item_label_print
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private UiFun As New UIfunc
    Const pageAbb As String = "gr_lbl_print"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        If Not IsPostBack Then
            ViewState("STOERE_CODE") = ""
            ViewState("DO_CODE") = ""

            ViewState("DO_CODE") = Server.UrlDecode(Request("DO_CODE"))
            ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

            BindGV()

        End If
    End Sub


    Private Sub BindGV()

        Dim DO_CODE, storer_code As String
        Dim dt As New DataTable

        DO_CODE = ViewState("DO_CODE")
        storer_code = ViewState("STORER_CODE")

        Dim sqlString As String = ""

        'sqlString = "SELECT WMS_DELV_ORDER_D.DOD_SEQ, WMS_ITEM.ITM_SKU_NO,WMS_DELV_ORDER_D.DOD_PACK_KEY, WMS_DELV_ORDER_D.DOD_PALLET_NO, WMS_DELV_ORDER_D.DOD_ITM_DESC, WMS_DELV_ORDER_D.DOD_BATCH_NO, WMS_ITEM.ITM_TYPE " & _
        '            "FROM WMS_DELV_ORDER_D INNER JOIN " & _
        '            "WMS_ITEM ON WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
        '            "WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
        '            " Where WMS_DELV_ORDER_D.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND WMS_DELV_ORDER_D.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
        '            " AND WMS_DELV_ORDER_D.DO_CODE='" & gU.dbEncode(DO_CODE) & "' "

        sqlString = " SELECT WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DO_PICKLIST_D.PLD_SERIAL_NO, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_DO_PICKLIST_D.PLD_SEQ, " & _
                    " WMS_ITEM.ITM_SKU_NO, WMS_DO_PICKLIST_D.PLD_LOC, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, " & _
                    " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_TYPE " & _
                    " FROM WMS_DO_PICKLIST_D INNER JOIN " & _
                    " WMS_ITEM ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                    " WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM.ITM_CODE AND WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                    " Where WMS_DO_PICKLIST_D.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND WMS_DO_PICKLIST_D.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
                    " AND WMS_DO_PICKLIST_D.DO_CODE='" & gU.dbEncode(DO_CODE) & "' "

        dt = gDB.getDataTable(sqlString)


        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()

    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("PLD_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_SEQ").ToString.Trim
                CType(e.Row.FindControl("ITM_TYPE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ITM_TYPE").ToString.Trim

        End Select
    End Sub

    Protected Sub SelectAll()
        If Not GridView1 Is Nothing And GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked = True
            Next
        End If
    End Sub

    Protected Sub UnselectAll()
        If Not GridView1 Is Nothing And GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked = False
            Next
        End If
    End Sub

    Protected Sub btnSelAll_Click(sender As Object, e As System.EventArgs) Handles btnSelAll.Click
        SelectAll()
    End Sub

    Protected Sub btnUnAll_Click(sender As Object, e As System.EventArgs) Handles btnUnAll.Click
        UnselectAll()
    End Sub

    Protected Sub btnPrint2_Click(sender As Object, e As System.EventArgs) Handles btnPrint2.Click
        PrintLabel()
    End Sub

    Protected Sub btnPrint_Click(sender As Object, e As System.EventArgs) Handles btnPrint.Click
        PrintLabel()
    End Sub

    Private Sub PrintLabel()

        Dim seqStr As String = ""
        Dim CableSeqStr As String = ""
        Dim DO_CODE As String = ViewState("DO_CODE")
        Dim storer_code As String = ViewState("STORER_CODE")

        Dim cableCount As Integer = 0
        Dim itemCount As Integer = 0

        If Not GridView1 Is Nothing And GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("nSelect"), CheckBox).Checked = True Then
                    'If CType(GridView1.Rows(i).FindControl("ITM_TYPE"), HiddenField).Value = "CABLE" Then
                    '    CableSeqStr &= CType(GridView1.Rows(i).FindControl("PLD_SEQ"), HiddenField).Value & "||"
                    '    cableCount += 1
                    'Else
                    seqStr &= CType(GridView1.Rows(i).FindControl("PLD_SEQ"), HiddenField).Value & "||"
                    itemCount += 1
                    ' End If

                End If
            Next
            If seqStr <> "" Then seqStr = Left(seqStr, Len(seqStr) - 2)
            'If CableSeqStr <> "" Then CableSeqStr = Left(CableSeqStr, Len(CableSeqStr) - 2)
        End If

        If cableCount + itemCount > 0 Then

            If itemCount > 0 Then ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "printLabel", "PrintLabel('" & storer_code & "', '" & seqStr & "', '" & DO_CODE & "');", True)
            'If cableCount > 0 Then ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "printLabelC", "PrintLabelCable('" & storer_code & "', '" & CableSeqStr & "', '" & DO_CODE & "');", True)
            'ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "closeWin", "window.close();", True)

        Else

            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "alertMsg", "alert('Please select item for printing.');", True)

        End If

    End Sub
End Class
