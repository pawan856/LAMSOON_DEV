Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class IMPORT_INTF_LOG_intferface_log
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then
            Call BindGV()
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
    End Sub

    Protected Sub BindGV()
        Dim selectSql As String
        Dim tmpDt As DataTable

        IMP_BATCH_ID.Text = Request("batch_id")

        If IMP_BATCH_ID.Text = "" Then
            Exit Sub
        End If


        selectSql = "select imp_batch_id, imp_file_name, itf_type, itf_imp_type, itf_status, itf_errcode, itf_errdesc, " & _
                        "convert(varchar, itf_start_dt, 20) as itf_start_dt, " & _
                        "convert(varchar, itf_end_dt, 20) as itf_end_dt " & _
                    "from wms_intf_log " & _
                    "where imp_batch_id = '" & gU.dbEncode(IMP_BATCH_ID.Text.Trim) & "' "

        tmpDt = gDB.getDataTable(selectSql)

        If tmpDt.Rows.Count > 0 Then
            IMP_FILE_NAME.Text = tmpDt.Rows(0).Item("IMP_FILE_NAME").ToString.Trim
            ITF_IMP_TYPE.Text = tmpDt.Rows(0).Item("ITF_IMP_TYPE").ToString.Trim
            ITF_START_DT.Text = tmpDt.Rows(0).Item("ITF_START_DT").ToString.Trim
            ITF_END_DT.Text = tmpDt.Rows(0).Item("ITF_END_DT").ToString.Trim
            ITF_STATUS.Text = tmpDt.Rows(0).Item("ITF_STATUS").ToString.Trim
            ITF_ERRCODE.Text = tmpDt.Rows(0).Item("ITF_ERRCODE").ToString.Trim
            ITF_ERRDESC.Text = tmpDt.Rows(0).Item("ITF_ERRDESC").ToString.Trim
        End If

        If ITF_IMP_TYPE.Text = "PO" Then
            trPO.Visible = True

            selectSql = "select * from wms_if_po_h " & _
                        "where imp_batch_id = '" & gU.dbEncode(IMP_BATCH_ID.Text.Trim) & "' " & _
                        "order by imp_rownum "

            tmpDt = gDB.getDataTable(selectSql)

            gvPO.DataSource = tmpDt
            gvPO.DataBind()

        ElseIf ITF_IMP_TYPE.Text = "IR" Then
            trIR.Visible = True

            selectSql = "select * from wms_if_ir_h " & _
                        "where imp_batch_id = '" & gU.dbEncode(IMP_BATCH_ID.Text.Trim) & "' " & _
                        "order by imp_rownum "

            tmpDt = gDB.getDataTable(selectSql)

            gvIR.DataSource = tmpDt
            gvIR.DataBind()

        ElseIf ITF_IMP_TYPE.Text = "TR" Then
            trTR.Visible = True

            selectSql = "select * from wms_if_tr_h " & _
                        "where imp_batch_id = '" & gU.dbEncode(IMP_BATCH_ID.Text.Trim) & "' " & _
                        "order by imp_rownum "

            tmpDt = gDB.getDataTable(selectSql)

            gvTR.DataSource = tmpDt
            gvTR.DataBind()

        End If

    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Then
            CType(ctl, Button).Enabled = True
            page_customizectrl = True
        End If
    End Function
End Class
