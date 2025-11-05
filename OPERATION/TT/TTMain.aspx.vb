Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class TTMain
    Inherits System.Web.UI.Page
    Private uiFun As New UIfunc
    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private codeValue As String

    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
        codeValue = BindFirstGV()
        Call BindGV(codeValue)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lblTitle.text = "Track and Trace"
            lblTTType.Text = "Type"
            lblTTNo.Text = "No."
            Submit.Text = "Search"
            gvrsList.EmptyDataText = "No Record Found."
        Else
            lblTitle.text = "Track and Trace"
            lblTTType.Text = "Type"
            lblTTNo.Text = "No."
            Submit.Text = "搜寻"
            gvrsList.EmptyDataText = "找不到相关资料"
        End If
        REM **********************

        If Not IsPostBack Then

        End If

        If Not IsPostBack Then
            If codeValue <> "" Then
                Call BindGV(codeValue)
                codeValue = BindFirstGV()
            End If
        End If

    End Sub

    Protected Sub gvrsList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvrsList.PageIndexChanging
        gvrsList.PageIndex = e.NewPageIndex
        Call BindGV(codeValue)
    End Sub

    Protected Sub gvrsList_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowCreated

        Select Case e.Row.RowType
            Case DataControlRowType.Header
                'Dim oGridView As GridView = DirectCast(sender, GridView)
                'Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here

                REM **********************

                'oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)

        End Select
    End Sub

    Protected Sub gvrsList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                CType(e.Row.FindControl("TTRO"), label).text = DataBinder.Eval(e.Row.DataItem, "RO").ToString.Trim
                CType(e.Row.FindControl("TTROGR_SB"), Image).ImageUrl = "../../images/TT/" & DataBinder.Eval(e.Row.DataItem, "ROGR").ToString.Trim & ".bmp"
                If DataBinder.Eval(e.Row.DataItem, "ROGR").ToString.Trim = "" Then
                    CType(e.Row.FindControl("TTROGR_SB"), image).visible = False
                End If
                CType(e.Row.FindControl("TTGR"), label).text = DataBinder.Eval(e.Row.DataItem, "GR").ToString.Trim
                CType(e.Row.FindControl("TTGRCO_SB"), Image).ImageUrl = "../../images/TT/" & DataBinder.Eval(e.Row.DataItem, "GRCO").ToString.Trim & ".bmp"
                If DataBinder.Eval(e.Row.DataItem, "GRCO").ToString.Trim = "" Then
                    CType(e.Row.FindControl("TTGRCO_SB"), image).visible = False
                End If
                CType(e.Row.FindControl("TTCO"), label).text = DataBinder.Eval(e.Row.DataItem, "CO").ToString.Trim
                CType(e.Row.FindControl("TTCODO_SB"), Image).ImageUrl = "../../images/TT/" & DataBinder.Eval(e.Row.DataItem, "CODO").ToString.Trim & ".bmp"
                If DataBinder.Eval(e.Row.DataItem, "CODO").ToString.Trim = "" Then
                    CType(e.Row.FindControl("TTCODO_SB"), image).visible = False
                End If
                CType(e.Row.FindControl("TTDO"), label).text = DataBinder.Eval(e.Row.DataItem, "DO").ToString.Trim

        End Select
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        DirectCast(sender, GridView).PageIndex = e.NewPageIndex

        BindFirstGV()

        'DirectCast(sender, GridView).DataBind()

        'Call BindGV(codeValue)
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here

                CType(e.Row.FindControl("gv1NO"), LinkButton).Text = DataBinder.Eval(e.Row.DataItem, "TTCODE").ToString.Trim
                CType(e.Row.FindControl("gv1NO"), LinkButton).CommandName = "UPDATE"
                '.Attributes.Add("href", "TTMain.aspx?codeValue=" & DataBinder.Eval(e.Row.DataItem, "TTCODE").ToString.Trim)
                CType(e.Row.FindControl("gv1TRACKNO"), label).Text = DataBinder.Eval(e.Row.DataItem, "TRACKNO").ToString.Trim
                CType(e.Row.FindControl("gv1CD"), label).Text = DataBinder.Eval(e.Row.DataItem, "TTDATE").ToString.Trim

        End Select
    End Sub

    Protected Function BindFirstGV() As String
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"

        REM **********************
        REM Modify Here

        Select Case TTType.SelectedValue
            Case "RO"
                SQLString = "SELECT RO_CODE as TTCODE, CONVERT(varchar(30), RO_DATE, 111) as TTDATE, RO_TRACK_NO as TRACKNO from WMS_REPLENISH"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " RO_CODE LIKE N'%" & TTNo.Text & "%' "
                End If
                lblsearchhd.Text = "RO"
            Case "GR"
                SQLString = "SELECT GR_CODE as TTCODE, CONVERT(varchar(30), GR_DATE, 111) as TTDATE, GR_TRACK_NO as TRACKNO  from WMS_GOODSRCV"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " GR_CODE LIKE N'%" & TTNo.Text & "%' "
                End If
                lblsearchhd.Text = "GR"
            Case "CO"
                SQLString = "SELECT CO_CODE as TTCODE, CONVERT(varchar(30), CO_DATE, 111) as TTDATE, CO_TRACK_NO as TRACKNO  from WMS_CUST_ORDER"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " CO_CODE LIKE N'%" & TTNo.Text & "%' "
                End If
                lblsearchhd.Text = "CO"
            Case "DO"
                SQLString = "SELECT DO_CODE as TTCODE, CONVERT(varchar(30), DO_DATE, 111) as TTDATE, DO_TRACK_NO as TRACKNO  from WMS_DELV_ORDER"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " DO_CODE LIKE N'%" & TTNo.Text & "%' "
                End If
                lblsearchhd.Text = "DO"
            Case "TRACKNO"
                SQLString = "SELECT RO_CODE as TTCODE, CONVERT(varchar(30), RO_DATE, 111) as TTDATE, RO_TRACK_NO as TRACKNO  from WMS_REPLENISH"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " RO_TRACK_NO LIKE N'%" & TTNo.Text & "%' "
                End If
                lblsearchhd.Text = "RO"
        End Select

        REM **********************

        If SCString <> "WHERE" Then
            SQLString = SQLString & " " & SCString
        End If
        dt = gDB.getDataTable(SQLString)
        If TTType.SelectedValue = "TRACKNO" Then
            If dt.Rows.Count = 0 Then
                SCString = "WHERE"
                SQLString = "SELECT CO_CODE as TTCODE, CONVERT(varchar(30), CO_DATE, 111) as TTDATE, CO_TRACK_NO as TRACKNO  from WMS_CUST_ORDER"
                If TTNo.Text <> "" Then
                    If SCString <> "WHERE" Then
                        SCString = SCString & " AND "
                    End If
                    SCString = SCString & " CO_TRACK_NO LIKE N'%" & TTNo.Text & "%' "
                End If
                If SCString <> "WHERE" Then
                    SQLString = SQLString & " " & SCString
                End If
                lblsearchhd.Text = "CO"
                dt = gDB.getDataTable(SQLString)
                If dt.Rows.Count = 0 Then
                    SQLString = "SELECT distinct WMS_DELV_ORDER.DO_CODE as TTCODE, CONVERT(varchar(30), WMS_DELV_ORDER.DO_DATE, 111) as TTDATE, WMS_DELV_ORDER_D.DOD_TRACK_NO as TRACKNO from WMS_DELV_ORDER, WMS_DELV_ORDER_D " & _
                    "WHERE WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE"
                    If TTNo.Text <> "" Then
                        SQLString = SQLString & " AND DOD_TRACK_NO LIKE N'%" & TTNo.Text & "%' "
                    End If
                    lblsearchhd.Text = "DO"
                    dt = gDB.getDataTable(SQLString)
                End If
            End If
        End If

        GridView1.DataSource = dt
        GridView1.DataBind()
        If dt.Rows.Count = 1 Then
            Return dt.Rows.Item(0)(0).ToString
        Else
            Return ""
        End If

    End Function

    Protected Sub BindGV(ByVal codeValue As String)
        Dim SQLString As String = ""
        Dim ROdt As New DataTable
        Dim GRdt As New DataTable
        Dim COdt As New DataTable
        Dim DOdt As New DataTable
        Dim newdt As New DataTable

        Dim SCString As String = "WHERE"

        REM **********************
        REM Modify Here

        Select Case TTType.SelectedValue
            Case "RO"
                SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE = '" & codeValue & "' "
                ROdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE from WMS_GOODSRCV WHERE GR_DOC_NO = '" & codeValue & "' AND GR_DOC_TYPE = 'RO' order by GR_CODE"
                GRdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE = '" & codeValue & "' order by CO_CODE "
                COdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CO_CODE IN (SELECT CO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE = '" & codeValue & "') order by DO_CO_CODE, DO_CODE"
                DOdt = gDB.getDataTable(SQLString)
            Case "GR"
                SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT GR_DOC_NO FROM WMS_GOODSRCV WHERE GR_CODE = '" & codeValue & "' AND GR_DOC_TYPE = 'RO') order by RO_CODE"
                ROdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE from WMS_GOODSRCV WHERE GR_CODE = '" & codeValue & "'"
                GRdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT GR_DOC_NO FROM WMS_GOODSRCV WHERE GR_CODE = '" & codeValue & "' AND GR_DOC_TYPE = 'RO')) order by CO_RO_CODE, CO_CODE"
                COdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CO_CODE IN (SELECT CO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT GR_DOC_NO FROM WMS_GOODSRCV WHERE GR_CODE = '" & codeValue & "' AND GR_DOC_TYPE = 'RO'))) order by DO_CO_CODE, DO_CODE"
                DOdt = gDB.getDataTable(SQLString)
            Case "CO"
                SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "')  order by RO_CODE"
                ROdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE from WMS_GOODSRCV WHERE GR_DOC_NO IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "')) AND GR_DOC_TYPE = 'RO' order by GR_DOC_NO, GR_CODE"
                GRdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "' "
                COdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CO_CODE = '" & codeValue & "' AND DO_CO_CODE <> '' order by DO_CODE"
                DOdt = gDB.getDataTable(SQLString)
            Case "DO"
                SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "')) order by RO_CODE"
                ROdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE  from WMS_GOODSRCV WHERE GR_DOC_NO IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "'))) AND GR_DOC_TYPE = 'RO' order by GR_DOC_NO, GR_CODE"
                GRdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "') order by CO_CODE, RO_CODE"
                COdt = gDB.getDataTable(SQLString)
                SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "'"
                DOdt = gDB.getDataTable(SQLString)
            Case "TRACKNO"
                If lblsearchhd.Text = "RO" Then
                    SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE = '" & codeValue & "' "
                    ROdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE from WMS_GOODSRCV WHERE GR_DOC_NO = '" & codeValue & "' AND GR_DOC_TYPE = 'RO' order by GR_CODE"
                    GRdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE = '" & codeValue & "' order by CO_CODE "
                    COdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CO_CODE IN (SELECT CO_CODE from WMS_CUST_ORDER WHERE CO_RO_CODE = '" & codeValue & "') order by DO_CO_CODE, DO_CODE"
                    DOdt = gDB.getDataTable(SQLString)
                ElseIf lblsearchhd.Text = "CO" Then
                    SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "')  order by RO_CODE"
                    ROdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE from WMS_GOODSRCV WHERE GR_DOC_NO IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "')) AND GR_DOC_TYPE = 'RO' order by GR_DOC_NO, GR_CODE"
                    GRdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_CODE = '" & codeValue & "' "
                    COdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CO_CODE = '" & codeValue & "' AND DO_CO_CODE <> '' order by DO_CODE"
                    DOdt = gDB.getDataTable(SQLString)
                Else
                    SQLString = "SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "')) order by RO_CODE"
                    ROdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT GR_CODE, GR_DOC_NO as RO_CODE  from WMS_GOODSRCV WHERE GR_DOC_NO IN (SELECT RO_CODE from WMS_REPLENISH WHERE RO_CODE IN (SELECT CO_RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "'))) AND GR_DOC_TYPE = 'RO' order by GR_DOC_NO, GR_CODE"
                    GRdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT CO_CODE, CO_RO_CODE as RO_CODE from WMS_CUST_ORDER WHERE CO_CODE IN (SELECT DO_CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "') order by CO_CODE, RO_CODE"
                    COdt = gDB.getDataTable(SQLString)
                    SQLString = "SELECT DO_CODE, DO_CO_CODE as CO_CODE from WMS_DELV_ORDER WHERE DO_CODE = '" & codeValue & "'"
                    DOdt = gDB.getDataTable(SQLString)
                End If
        End Select

        REM **********************
        REM create New DT
        newdt.Columns.Add("RO")
        newdt.Columns.Add("ROGR")
        newdt.Columns.Add("GR")
        newdt.Columns.Add("GRCO")
        newdt.Columns.Add("CO")
        newdt.Columns.Add("CODO")
        newdt.Columns.Add("DO")
        REM ***************

        Dim ROIDT As Integer
        ROIDT = 0
        Dim COIDT As Integer
        COIDT = 0

        For roi = 0 To ROdt.Rows.Count
            If ROdt.Rows.Count > 0 And ROdt.Rows.Count = roi Then
                Exit For
            End If
            REM RO
            If ROdt.Rows.Count > 0 Then
                Call addDataRow(newdt, ROIDT)
                newdt.Rows.Item(ROIDT)("RO") = ROdt.Rows.Item(roi)(0)
            End If
            REM GR
            For gri = 0 To GRdt.Rows.Count - 1
                If ROdt.Rows.Count > 0 Then
                    If GRdt.Rows.Item(gri)(1).ToString = ROdt.Rows.Item(roi)(0).ToString Then

                    Else
                        ROIDT = gri
                    End If
                End If
                Call addDataRow(newdt, gri)
                newdt.Rows.Item(gri)("GR") = GRdt.Rows.Item(gri)(0)
                If GRdt.Rows.Count > 1 Then
                    If gri = 0 Then
                        newdt.Rows.Item(gri)("ROGR") = "sd"
                    ElseIf GRdt.Rows.Count - 1 = gri Then
                        newdt.Rows.Item(gri)("ROGR") = "lbottom"
                    ElseIf GRdt.Rows.Count - 1 > gri Then
                        newdt.Rows.Item(gri)("ROGR") = "lc"
                    End If
                Else
                    If ROdt.Rows.Count > 0 Then
                        newdt.Rows.Item(gri)("ROGR") = "ch"
                    End If
                End If
            Next
            REM CO
            For coi = 0 To COdt.Rows.Count
                If COIDT < coi Then
                    COIDT = coi
                End If
                If COdt.Rows.Count > 0 And COdt.Rows.Count = coi Then
                    Exit For
                End If
                If ROdt.Rows.Count > 0 Then
                    If COdt.Rows.Count > 0 Then
                        If COdt.Rows.Item(coi)(1).ToString = ROdt.Rows.Item(roi)(0).ToString Then

                        Else
                            ROIDT = coi
                        End If
                    End If
                End If
                If COdt.Rows.Count > 0 Then
                    Call addDataRow(newdt, COIDT)
                    newdt.Rows.Item(COIDT)("CO") = COdt.Rows.Item(coi)(0)

                    If COdt.Rows.Count > 1 Then
                        If coi = 0 Then
                            newdt.Rows.Item(COIDT)("GRCO") = "sd"
                        ElseIf COdt.Rows.Count - 1 = coi Then
                            newdt.Rows.Item(COIDT)("GRCO") = "rb"
                        ElseIf COdt.Rows.Count - 1 > coi Then
                            newdt.Rows.Item(COIDT)("GRCO") = "rc"
                        End If
                    Else
                        If ROdt.Rows.Count > 0 Then
                            newdt.Rows.Item(COIDT)("GRCO") = "ch"
                        End If
                    End If
                End If

                REM DO
                Dim newLink As Boolean = False
                For doi = 0 To DOdt.Rows.Count - 1
                    If COdt.Rows.Count > 0 Then
                        If DOdt.Rows.Item(doi)(1).ToString = COdt.Rows.Item(coi)(0).ToString Then
                            Call addDataRow(newdt, COIDT)
                            newdt.Rows.Item(COIDT)("DO") = DOdt.Rows.Item(doi)(0)

                            If DOdt.Rows.Count > doi + 1 Then
                                If DOdt.Rows.Item(doi + 1)(1).ToString <> COdt.Rows.Item(coi)(0).ToString Then
                                    newLink = True
                                End If
                            End If

                            If DOdt.Rows.Count > 1 Then
                                If doi = 0 Then
                                    newdt.Rows.Item(COIDT)("CODO") = "sd"
                                ElseIf DOdt.Rows.Count - 1 = doi Then

                                    newdt.Rows.Item(COIDT)("CODO") = "rb"

                                ElseIf DOdt.Rows.Count - 1 > doi Then
                                    If newLink = True Then
                                        newdt.Rows.Item(COIDT)("CODO") = "rb"
                                        newLink = False
                                    Else
                                        If doi - 1 >= 0 Then
                                            If DOdt.Rows.Item(doi - 1)(1).ToString <> COdt.Rows.Item(coi)(0).ToString Then
                                                newdt.Rows.Item(COIDT)("CODO") = "sd"
                                            Else
                                                newdt.Rows.Item(COIDT)("CODO") = "rc"
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If COdt.Rows.Count > 0 Then
                                    newdt.Rows.Item(COIDT)("CODO") = "ch"
                                End If
                            End If
                            COIDT = COIDT + 1
                            If COdt.Rows.Count - 1 > coi Then
                                If COIDT > coi Then
                                    Call addDataRow(newdt, COIDT)
                                    newdt.Rows.Item(COIDT)("GRCO") = "c"
                                End If
                            End If
                        End If
                    Else
                        Call addDataRow(newdt, doi)
                        newdt.Rows.Item(doi)("DO") = DOdt.Rows.Item(doi)(0)
                        If DOdt.Rows.Count > 1 Then
                            If doi = 0 Then
                                newdt.Rows.Item(doi)("CODO") = "sd"
                            ElseIf DOdt.Rows.Count - 1 = doi Then
                                newdt.Rows.Item(doi)("CODO") = "rb"
                            ElseIf DOdt.Rows.Count - 1 > doi Then
                                newdt.Rows.Item(doi)("CODO") = "rc"
                            End If
                        Else
                            If COdt.Rows.Count > 0 Then
                                newdt.Rows.Item(doi)("CODO") = "ch"
                            End If
                        End If
                    End If
                Next
            Next
            If ROIDT < roi Then
                ROIDT = roi
            End If
        Next

        'Select Case TTType.SelectedValue
        '    Case "RO"
        '        For i = 0 To GRdt.Rows.Count - 1
        '            If GRdt.Rows.Item(i)(0) <> "" Then
        '                If GRdt.Rows.Count > 1 Then
        '                    If i = 0 Then
        '                        newdt.Rows.Item(i)("ROGR") = "sd"
        '                    Else
        '                        If GRdt.Rows.Count > 2 Then
        '                            If i = GRdt.Rows.Count - 1 Then
        '                                newdt.Rows.Item(i)("ROGR") = "rb"
        '                            Else
        '                                newdt.Rows.Item(i)("ROGR") = "c"
        '                            End If
        '                        Else
        '                            newdt.Rows.Item(i)("ROGR") = "rb"
        '                        End If
        '                    End If
        '                ElseIf GRdt.Rows.Count = 1 Then
        '                    newdt.Rows.Item(i)("ROGR") = "ch"
        '                Else
        '                    newdt.Rows.Item(i)("ROGR") = ""
        '                End If
        '            Else
        '                newdt.Rows.Item(i)("ROGR") = ""
        '            End If
        '        Next
        '    Case "GR"
        '        For i = 0 To ROdt.Rows.Count - 1
        '            If ROdt.Rows.Item(i)(0) <> "" Then
        '                If ROdt.Rows.Count > 1 Then
        '                    If i = 0 Then
        '                        newdt.Rows.Item(i)("ROGR") = "sd"
        '                    Else
        '                        If ROdt.Rows.Count > 2 Then
        '                            If i = COdt.Rows.Count - 1 Then
        '                                newdt.Rows.Item(i)("ROGR") = "lbottom"
        '                            Else
        '                                newdt.Rows.Item(i)("ROGR") = "c"
        '                            End If
        '                        Else
        '                            newdt.Rows.Item(i)("ROGR") = "lbottom"
        '                        End If
        '                    End If
        '                ElseIf ROdt.Rows.Count = 1 Then
        '                    newdt.Rows.Item(i)("ROGR") = "ch"
        '                Else
        '                    newdt.Rows.Item(i)("ROGR") = ""
        '                End If
        '            Else
        '                newdt.Rows.Item(i)("ROGR") = ""
        '            End If
        '        Next
        '    Case "CO"
        '        For i = 0 To DOdt.Rows.Count - 1
        '            If DOdt.Rows.Item(i)(0) <> "" Then
        '                If DOdt.Rows.Count > 1 Then
        '                    If i = 0 Then
        '                        newdt.Rows.Item(i)("CODO") = "sd"
        '                    Else
        '                        If DOdt.Rows.Count > 2 Then
        '                            If i = DOdt.Rows.Count - 1 Then
        '                                newdt.Rows.Item(i)("CODO") = "rb"
        '                            Else
        '                                newdt.Rows.Item(i)("CODO") = "c"
        '                            End If
        '                        Else
        '                            newdt.Rows.Item(i)("CODO") = "rb"
        '                        End If
        '                    End If
        '                ElseIf DOdt.Rows.Count = 1 Then
        '                    newdt.Rows.Item(i)("CODO") = "ch"
        '                Else
        '                    newdt.Rows.Item(i)("CODO") = ""
        '                End If
        '            Else
        '                newdt.Rows.Item(i)("CODO") = ""
        '            End If
        '        Next
        '    Case "DO"
        '        For i = 0 To COdt.Rows.Count - 1
        '            If COdt.Rows.Item(i)(0) <> "" Then
        '                If COdt.Rows.Count > 1 Then
        '                    If i = 0 Then
        '                        newdt.Rows.Item(i)("CODO") = "sd"
        '                    Else
        '                        If COdt.Rows.Count > 2 Then
        '                            If i = COdt.Rows.Count - 1 Then
        '                                newdt.Rows.Item(i)("CODO") = "lbottom"
        '                            Else
        '                                newdt.Rows.Item(i)("CODO") = "c"
        '                            End If
        '                        Else
        '                            newdt.Rows.Item(i)("CODO") = "lbottom"
        '                        End If
        '                    End If
        '                ElseIf COdt.Rows.Count = 1 Then
        '                    newdt.Rows.Item(i)("CODO") = "ch"
        '                Else
        '                    newdt.Rows.Item(i)("CODO") = ""
        '                End If
        '            Else
        '                newdt.Rows.Item(i)("CODO") = ""
        '            End If
        '        Next
        'End Select

        gvrsList.DataSource = newdt
        gvrsList.DataBind()


        Select Case TTType.SelectedValue
            Case "RO"
                For i = 0 To gvrsList.Rows.Count - 1
                    CType(gvrsList.Rows(i).FindControl("TTRO"), Label).Font.Bold = True
                Next
            Case "GR"
                For i = 0 To gvrsList.Rows.Count - 1
                    CType(gvrsList.Rows(i).FindControl("TTGR"), Label).Font.Bold = True
                Next
            Case "CO"
                For i = 0 To gvrsList.Rows.Count - 1
                    CType(gvrsList.Rows(i).FindControl("TTCO"), Label).Font.Bold = True
                Next
            Case "DO"
                For i = 0 To gvrsList.Rows.Count - 1
                    CType(gvrsList.Rows(i).FindControl("TTDO"), Label).Font.Bold = True
                Next
            Case "TRACKNO"
                For i = 0 To gvrsList.Rows.Count - 1
                    If lblsearchhd.Text = "RO" Then
                        CType(gvrsList.Rows(i).FindControl("TTRO"), Label).Font.Bold = True
                    ElseIf lblsearchhd.Text = "CO" Then
                        CType(gvrsList.Rows(i).FindControl("TTCO"), Label).Font.Bold = True
                    Else
                        CType(gvrsList.Rows(i).FindControl("TTDO"), Label).Font.Bold = True
                    End If
                Next
        End Select
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If TTNO.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "No. cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "No. 不能空白!", Session("gLang"))
            End If
            Return False
        End If
        Return True

    End Function

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        codeValue = CType(GridView1.Rows(e.RowIndex).FindControl("gv1NO"), LinkButton).Text
        Call BindGV(codeValue)
    End Sub

    Protected Sub addDataRow(ByRef newdt As DataTable, ByVal checkNo As Integer)
        For addDRi = 0 To checkNo - newdt.Rows.Count
            Dim newrows As DataRow = newdt.Rows.Add
        Next
    End Sub


    Protected Sub TTType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TTType.SelectedIndexChanged
        codeValue = BindFirstGV()
        Call BindGV(codeValue)
    End Sub
End Class
