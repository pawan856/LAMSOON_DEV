Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports System.Data.SqlClient
Partial Class main_news
    Inherits System.Web.UI.Page

    Private ar As New AccessRightUtils
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private moduleAction As String = ""
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Not IsPostBack Then
            BindGV()
        End If


    End Sub
    Private Sub BindGV()
        Dim sqlString As String = ""
        Dim rhtList As String = ""
        Dim dt As DataTable

        sqlString = "SELECT GRP_CODE FROM WMS_USER_GROUP_ALLOC " & _
                    "where USR_ID='" & Session("usr_id") & "' and GRP_CODE in ('ISSUE_GP','LAMMA_GP','RECEIVING_GP','ADMIN_GP','MANAGEMENT_GP')"

        dt = gDB.getDataTable(sqlString)

        If dt.Rows.Count > 0 Then

            For i = 0 To dt.Rows.Count - 1
                rhtList = gU.appendToList(rhtList, dt.Rows(i).Item("GRP_CODE").ToString.Trim)
            Next

            If Not String.IsNullOrWhiteSpace(rhtList) Then
                PO_ALC_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL("select count(distinct RO_CODE) as RO_COUNT_ALC from WMS_REPLENISH where RO_STATUS NOT IN('CLOSED', 'CANCELLED') and imp_code='" & Session("imp_code") & "'"), 0)
                SR_ALC_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL("select count(distinct RT_CODE) as SR_COUNT_ALC from WMS_STOCK_RETURN where RT_STATUS NOT IN('CLOSED', 'POSTED', 'CANCELLED') and imp_code='" & Session("imp_code") & "'"), 0)

                sqlString = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SELECT count(distinct WMS_CUST_ORDER.CO_CODE) as CO_COUNT_ALC " &
                            "FROM WMS_CUST_ORDER INNER JOIN WMS_CUST_ORDER_D ON WMS_CUST_ORDER_D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE AND WMS_CUST_ORDER_D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE AND WMS_CUST_ORDER_D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                            "WHERE CO_STATUS NOT IN('CLOSED', 'CANCELLED') AND WMS_CUST_ORDER.imp_code='" & Session("imp_code") & "'"
                SIR_ALC_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

                sqlString = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SELECT count(distinct WMS_DELV_ORDER.DO_CODE) as DO_COUNT_ALC " &
                            "FROM WMS_DELV_ORDER INNER JOIN WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE " &
                            "WHERE DO_STATUS NOT IN('POSTED', 'CANCELLED') AND WMS_DELV_ORDER.imp_code='" & Session("imp_code") & "'"
                WIT_ALC_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

                PO_LMA_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL("select count(distinct RO_CODE) as RO_COUNT_LMA from WMS_REPLENISH where RO_STATUS NOT IN('CLOSED', 'CANCELLED') and RO_WH_CODE='LMA' and imp_code='" & Session("imp_code") & "'"), 0)
                SR_LMA_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL("select count(distinct RT_CODE) as SR_COUNT_ALC from WMS_STOCK_RETURN where RT_STATUS NOT IN('CLOSED', 'POSTED', 'CANCELLED') and RT_WH='LMA' and imp_code='" & Session("imp_code") & "'"), 0)

                sqlString = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SELECT count(distinct WMS_CUST_ORDER.CO_CODE) as CO_COUNT_LMA " &
                            "FROM WMS_CUST_ORDER INNER JOIN WMS_CUST_ORDER_D ON WMS_CUST_ORDER_D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE AND WMS_CUST_ORDER_D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE AND WMS_CUST_ORDER_D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                            "WHERE CO_STATUS NOT IN('CLOSED', 'CANCELLED') AND WMS_CUST_ORDER_D.COD_WH_CODE='LMA' and WMS_CUST_ORDER.imp_code='" & Session("imp_code") & "'"
                SIR_LMA_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

                sqlString = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SELECT count(distinct WMS_DELV_ORDER.DO_CODE) as DO_COUNT_LMA " &
                            "FROM WMS_DELV_ORDER INNER JOIN WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE " &
                            "WHERE DO_STATUS NOT IN('POSTED', 'CANCELLED') AND WMS_DELV_ORDER_D.DOD_WH_CODE='LMA'  and WMS_DELV_ORDER.imp_code='" & Session("imp_code") & "'"
                WIT_LMA_COUNT.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

                If gU.inList(rhtList, "ADMIN_GP") OrElse gU.inList(rhtList, "MANAGEMENT_GP") Then
                    RCV_PNL.Visible = True
                    ISS_PNL.Visible = True
                    'LMA_PNL.Visible = True

                Else
                    If gU.inList(rhtList, "RECEIVING_GP") Then
                        RCV_PNL.Visible = True
                    Else
                        RCV_PNL.Visible = False
                    End If

                    If gU.inList(rhtList, "ISSUE_GP") Then
                        ISS_PNL.Visible = True
                    Else
                        ISS_PNL.Visible = False
                    End If

                    If gU.inList(rhtList, "LAMMA_GP") Then
                        'LMA_PNL.Visible = True
                    Else
                        LMA_PNL.Visible = False
                    End If

                End If

                emptyTR.Visible = False
                ToDo_POPUP.Show()
            End If


        Else
            emptyTR.Visible = True
            pnlHeader.Visible = False
            'btnToDo.Visible = False
        End If

        sqlString = "SELECT COUNT(*) from wms_alert where ALRT_TO_USERID='" & gU.dbEncode(Session("usr_id")) & "'"
        Dim alertCount As Integer = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

        Select Case ("glang")
            Case "C"
                alertLink.Text = "你有 " & alertCount & " 個工單題示."
            Case Else
                alertLink.Text = "You have " & alertCount & " Note Alert(s)."
        End Select
        ToDo_POPUP.Show()

    End Sub

    Protected Sub BTN_PO_ALC_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_PO_ALC_COUNT.Click
        GoToSearch("RO", "ALC")
    End Sub

    Protected Sub BTN_PO_LMA_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_PO_LMA_COUNT.Click
        GoToSearch("RO", "LMA")
    End Sub

    Protected Sub BTN_SIR_ALC_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_SIR_ALC_COUNT.Click
        GoToSearch("CO", "ALC")
    End Sub

    Protected Sub BTN_SIR_LMA_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_SIR_LMA_COUNT.Click
        GoToSearch("CO", "LMA")
    End Sub

    Protected Sub BTN_SR_ALC_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_SR_ALC_COUNT.Click
        GoToSearch("SR", "ALC")
    End Sub

    Protected Sub BTN_SR_LMA_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_SR_LMA_COUNT.Click
        GoToSearch("SR", "LMA")
    End Sub

    Protected Sub BTN_WIT_ALC_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_WIT_ALC_COUNT.Click
        GoToSearch("DO", "ALC")
    End Sub

    Protected Sub BTN_WIT_LMA_COUNT_Click(sender As Object, e As System.EventArgs) Handles BTN_WIT_LMA_COUNT.Click
        GoToSearch("DO", "LMA")
    End Sub

    Private Sub GoToSearch(ByVal iType As String, ByVal cMWH As String)
        Dim fun_code As String = ""

        Select Case iType
            Case "RO"
                fun_code = "IB_RO"
            Case "SR"
                fun_code = "IB_SR"
            Case "CO"
                fun_code = "OB_CO"
            Case "DO"
                fun_code = "OB_DO"

        End Select

        Session("mWH") = cMWH

        Dim rmtPost As New RemotePost
        rmtPost.Url = "cms_search.aspx"
        rmtPost.Add("menu_code", fun_code)
        rmtPost.Add("sl", True)
        'rmtPost.Add("cMWH", cMWH)
        rmtPost.Post()

    End Sub

    Protected Sub btnToDo_Click(sender As Object, e As System.EventArgs) Handles btnToDo.Click
        BindGV()
    End Sub
End Class
