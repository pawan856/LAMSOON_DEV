Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Globalization

Partial Class INBOUND_GR_GRInsp
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private uiFun As New UIfunc
    Private cU As New CommonUtils
    Private dtRejReason, dtDtlStatus, dtInspBy As DataTable

    Dim DDFORMAT As String = gU.getConfig("DDFORMATNO")

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If


        If Not IsPostBack Then
            BindGV()
        End If

    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(btnSave)

        If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
            'Set labels, attributes, etc
            setGeneralControls()

            'Set field access (hide, readonly, view mode, etc)
            setPageCtrlAccess()
        End If

        moduleAction.Value = ""
    End Sub

    Private Sub setGeneralControls()
        lheader.Text = "Inspection"
    End Sub

    Private Sub setPageCtrlAccess()
        Dim exceptionEditList As List(Of String)

        exceptionEditList = New List(Of String)

        'exceptionEditList.Add("btnInsp")

        If GR_STATUS.Value = "CANCELLED" Then
            ar.sec_write = "N"
            btnSave.Visible = False
            'ElseIf GR_STATUS.Value = "POSTED" Then
            '    ar.sec_write = "N"
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

    End Sub

    Private Sub BindGV()
        Dim selectSql As String
        Dim insp_dt As DataTable

        IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
        STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
        GR_CODE.Value = Server.UrlDecode(Request("GR_CODE"))
        GR_STATUS.Value = Server.UrlDecode(Request("GR_STATUS"))

        'insp_dt = Session("gr_insp_dt")

        selectSql = "select i.IMP_CODE, i.STORER_CODE, i.GR_CODE, i.GRD_SEQ, i.GRI_SEQ, i.GRI_ITM_CODE, i.GRI_PACK_KEY, i.GRI_PALLET_NO, i.GRI_BATCH_NO, i.GRI_SERIAL_NO_LIST, i.GRI_INSP_QTY, " & _
                       "i.GRI_PASS_QTY, i.GRI_REJ_QTY, i.GRI_REJ_REASON, i.GRI_PHOTO, i.GRI_STATUS, i.GRI_REMARK, i.GRI_INSPECTED_BY, " & _
                       "t.ITM_NAME, t.ITM_SKU_NO, " & _
                       "'U' as mFlag " & _
                   "from WMS_GOODSRCV_INSP i " & _
                   "left outer join WMS_ITEM t " & _
                       "on i.IMP_CODE = t.IMP_CODE " & _
                       "and i.STORER_CODE = t.STORER_CODE " & _
                       "and i.GRI_ITM_CODE = t.ITM_CODE " & _
                       "and i.GRI_PACK_KEY = t.PACK_KEY " & _
                   "where i.gr_code = '" & gU.dbEncode(GR_CODE.Value) & "' " & _
                   "and i.storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                   "and i.imp_code = '" & Session("imp_code") & "' " & _
                   "order by i.GRI_ITM_CODE, i.GRI_PACK_KEY, i.GRI_BATCH_NO, i.GRI_SEQ "

        insp_dt = gDB.getDataTable(selectSql)

        Session("gr_insp_dt") = insp_dt

        GridView1.DataSource = insp_dt

        GridView1.DataBind()
    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False

        If ctl.ID = "ser_table" Then
            customizectrl = True
        End If
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Or ctl.ID = "btnPutAway" Or ctl.ID = "btnUnPost" Then
            If ctl.ID = "btnUnpost" Then
                CType(ctl, Button).Enabled = ar.hasBtnRight("BT_GR_UNPOST")
            Else
                CType(ctl, Button).Enabled = True
            End If

            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        Dim selectSql As String

        selectSql = "select colc_code as code, colc_eng_value as name " & _
                    "from wms_col_code " & _
                    "where colc_tabcol = 'WMS_GOODSRCV_D.GRD_REJ_REASON' " & _
                    "order by colc_eng_value "

        dtRejReason = gDB.getDataTable(selectSql)

        selectSql = "select colc_code as code, colc_eng_value as name from wms_col_code where colc_tabcol = 'WMS_GOODSRCV_INSP.GRI_STATUS' order by COLC_DISPLAY_SEQ"

        dtDtlStatus = gDB.getDataTable(selectSql)

        selectSql = "select usr_id as code, usr_fname + ' ' + usr_sname as name " & _
                    "from wms_user " & _
                    "where usr_status = 'A' " & _
                    "and usr_type = 'S' "

        dtInspBy = gDB.getDataTable(selectSql)
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("gri_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "gri_itm_code").ToString.Trim
                CType(e.Row.FindControl("grd_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "grd_seq").ToString.Trim
                CType(e.Row.FindControl("gri_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "gri_seq").ToString.Trim

                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim

                CType(e.Row.FindControl("gri_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "gri_pack_key").ToString.Trim
                CType(e.Row.FindControl("gri_pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "gri_pallet_no").ToString.Trim
                CType(e.Row.FindControl("gri_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "gri_batch_no").ToString.Trim

                CType(e.Row.FindControl("gri_insp_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "gri_insp_qty").ToString.Trim
                CType(e.Row.FindControl("gri_pass_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "gri_pass_qty").ToString.Trim
                CType(e.Row.FindControl("gri_rej_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "gri_rej_qty").ToString.Trim

                CType(e.Row.FindControl("gri_remark"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "gri_remark").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("gri_rej_reason"), DropDownList), dtRejReason, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "gri_rej_reason").ToString.Trim)

                uiFun.load_dropdown(CType(e.Row.FindControl("gri_status"), DropDownList), dtDtlStatus, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "gri_status").ToString.Trim, True)

                uiFun.load_dropdown(CType(e.Row.FindControl("gri_inspected_by"), DropDownList), dtInspBy, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "gri_inspected_by").ToString.Trim)

        End Select
    End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim i As Integer

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("gri_pass_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Pass Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 合格數量!", Session("gLang"))
                        End If
                        CType(GridView1.Rows(i).FindControl("gri_pass_qty"), TextBox).Focus()
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("gri_rej_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Rejected Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 拒收數量!", Session("gLang"))
                        End If
                        CType(GridView1.Rows(i).FindControl("gri_rej_qty"), TextBox).Focus()
                        Return False
                    End If

                    If CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("gri_pass_qty"), TextBox).Text.Trim, "0")) + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("gri_rej_qty"), TextBox).Text.Trim, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("gri_insp_qty"), Label).Text.Trim, "0")) Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Pass Qty + Rejected Qty cannot greater than Insp. Qty!", Session("gLang"))
                        CType(GridView1.Rows(i).FindControl("gri_pass_qty"), TextBox).Focus()
                        Return False
                    End If


                End If
            Next
        End If

        Return True
    End Function

    Private Sub save()
        Dim cnn As SqlConnection
        Dim transaction As SqlTransaction
        Dim insp_dt As DataTable
        Dim updateSql As String

        cnn = gDB.getConnection()
        transaction = cnn.BeginTransaction()

        Try
            insp_dt = Session("gr_insp_dt")

            If cU.gfBuildDataTableforGridView(insp_dt, GridView1, True) Then

                For Each rows As DataRow In insp_dt.Rows
                    Select Case rows.Item("mFlag")
                        Case "U"

                            updateSql = "update WMS_GOODSRCV_INSP " & _
                                        "set gri_pass_qty = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gri_pass_qty").ToString.Trim, "0")) & "', " & _
                                            "gri_rej_qty = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("gri_rej_qty").ToString.Trim, "0")) & "', " & _
                                            "gri_rej_reason = '" & gU.dbEncode(gU.decodeNull(rows.Item("gri_rej_reason").ToString.Trim, "")) & "', " & _
                                            "gri_status = '" & gU.dbEncode(gU.decodeNull(rows.Item("gri_status").ToString.Trim, "")) & "', " & _
                                            "gri_remark = N'" & gU.dbEncode(gU.decodeNull(rows.Item("gri_remark").ToString.Trim, "")) & "', " & _
                                            "gri_inspected_by = '" & gU.dbEncode(gU.decodeNull(rows.Item("gri_inspected_by").ToString.Trim, "")) & "', " & _
                                            "sys_lub = '" & Session("usr_id") & "', " & _
                                            "sys_lud = Getdate() " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                                        "and GR_CODE = '" & gU.dbEncode(GR_CODE.Value) & "' " & _
                                        "and GRD_SEQ = '" & rows.Item("GRD_SEQ").ToString.Trim & "' " & _
                                        "and GRI_SEQ = '" & rows.Item("GRI_SEQ").ToString.Trim & "' "

                            gDB.amendData(updateSql, cnn, transaction)
                    End Select
                Next

            End If

            transaction.Commit()

            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "SAVE_OK", "alert('Inspection record(s) has been saved!');window.close();", True)

        Catch ex As Exception
            'transaction.Rollback()
            'Response.Write(ex.Message)
            'uiFun.displayMsgNew(updtPnlAlert, "1008", "", Session("gLang"))

            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

            If Not ScriptManager.GetCurrent(Me).IsInAsyncPostBack Then
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT1", "Record save failure!", Session("gLang"))
                uiFun.displayMsgNew(updtPnlAlert, "EX_ALERT2", ex.Message, Session("gLang"))
            End If

        End Try
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click
        If validateAll() Then
            save()
        End If
    End Sub
End Class
