Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class LOOKUP_locMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private db As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private thumb As ThumbGenerator

    Private pItemArray As String()
    Private newDataArray As String()
    Private lWH_CODE As String = ""
    Private lFL_NUM As String = ""

    Private flDt As New DataTable
    Private arDt As New DataTable
    Private rkDt As New DataTable
    Private bnsaveDt As New DataTable
    Private bnDt As New DataTable

    Private arDict As New Dictionary(Of String, DataRow)
    Private rkDict As New Dictionary(Of String, DataRow)
    Private bnDict As New Dictionary(Of String, DataRow)

    Private flPhoto As New Dictionary(Of String, HttpPostedFile)
    Private arPhoto As New Dictionary(Of String, HttpPostedFile)

    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("MAST_WM", Session("usr_id"), Me)
        ar.hideForm(Me)

        If Request("pForm") <> "" Then
            pForm.Value = Request("pForm")
        End If
        If Request("pItemList") <> "" Then
            pItemList.Value = Request("pItemList")
        End If

        pItemArray = Split(pItemList.Value, ", ")
        newDataArray = Split(pItemList.Value, ", ")
        If Request("wh") <> "" Then
            wh.Value = Request("wh")
            WH_CODE.Value = Request("wh")
        Else
            Response.Write("missing warehouse")
            Exit Sub
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            Submit.Text = "Select"
        Else
            Submit.Text = "Select"
        End If
        REM **********************

        If Not IsPostBack Then
            'Session("arDict") = Nothing
            'Session("rkDict") = Nothing
            'Session("bnDict") = Nothing
            'Session("flPhoto") = Nothing
            'Session("arPhoto") = Nothing

            ViewState("pre_flcode") = ""
            ViewState("pre_arcode") = ""
            ViewState("pre_rkcode") = ""
            ViewState("pre_bncode") = ""

            ViewState("first_load") = True
            ViewState("loadbn") = True
            ViewState("flDt") = Nothing
            ViewState("arDt") = Nothing
            ViewState("rkDt") = Nothing
            ViewState("bnsaveDt") = Nothing
            ViewState("bnDt") = Nothing

            'fillingUpTables()
        Else
            'If Not Session("arDict") Is Nothing Then arDict = Session("arDict")
            'If Not Session("rkDict") Is Nothing Then rkDict = Session("rkDict")
            'If Not Session("bnDict") Is Nothing Then bnDict = Session("bnDict")
            'If Not Session("flPhoto") Is Nothing Then flPhoto = Session("flPhoto")
            'If Not Session("arPhoto") Is Nothing Then arPhoto = Session("arPhoto")

            flDt = ViewState("flDt")
            arDt = ViewState("arDt")
            rkDt = ViewState("rkDt")
            bnsaveDt = ViewState("bnsaveDt")
        End If

        If Session("gLang") = "E" Then
            lbl_FL_NUM.Text = "Floor:"
            lbl_FL_NAME.Text = "Floor Name:"
            lbl_FL_NAME_CH.Text = "Floor Name in Chinese:"
            lbl_FL_GROSS_AREA.Text = "Gross Area:"
            lbl_FL_GROSS_CBM.Text = "Gross CBM:"
            lbl_FL_MAX_WEIGHT.Text = "Max Weight:"
            lbl_FL_LENGTH.Text = "Length:"
            lbl_FL_WIDTH.Text = "Width:"
            lbl_FL_HEIGHT.Text = "Height:"
            'lbl_FL_PICTURE.text = 

            lbl_AR_CODE.Text = "Area Code:"
            lbl_AR_NAME.Text = "Area Name:"
            lbl_AR_NAME_CH.Text = "Area Name in Chinese:"
            lbl_AR_X.Text = "Columns:"
            lbl_AR_Y.Text = "Rows:"
            lbl_AR_LENGTH.Text = "Length:"
            lbl_AR_WIDTH.Text = "Width:"
            lbl_AR_HEIGHT.Text = "Height:"
            lbl_AR_GROSS_AREA.Text = "Gross Area:"
            lbl_AR_GROSS_CBM.Text = "Gross CBM:"
            lbl_AR_FACILITIES.Text = "Facilities:"
            'lbl_AR_PICTURE.text = 

            lbl_RK_CODE.Text = "Rack Code:"
            lbl_RK_NAME.Text = "Rack Name:"
            lbl_RK_NAME_CH.Text = "Rack Name in Chinese:"
            lbl_RK_GROSS_AREA.Text = "Gross Area:"
            lbl_RK_GROSS_CBM.Text = "Gross CBM"
            lbl_RK_X.Text = "Rack Columns:"
            lbl_RK_Y.Text = "Rack Rows:"
            lbl_RK_LENGTH.Text = "Length:"
            lbl_RK_WIDTH.Text = "Width:"
            lbl_RK_DEPTH.Text = "Depth:"
            lbl_RK_XBINS.Text = "Bins Columns:"
            lbl_RK_YBINS.Text = "Bins Rows:"
            lbl_RK_REM.Text = "Remarks:"

            lbl_BN_CODE.Text = "Bins Code:"
            lbl_BN_X.Text = "Bin Columns No.:"
            lbl_BN_Y.Text = "Bin Rows No.:"
            lbl_BN_LENGTH.Text = "Length:"
            lbl_BN_WIDTH.Text = "Width:"
            lbl_BN_DEPTH.Text = "Depth:"
            lbl_BN_REM.Text = "Remarks:"
            lbl_BN_STATUS.Text = "Status:"
            lbl_BN_CBM.Text = "CBM:"
            lbl_BN_CSMS_CODE.Text = "EDI Code"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"

            saveBtn1.OnClientClick = "if (confirm(""Are you sure to save this record?"")) {window.document.forms[0].mode.value=""SAVE"";}else{return false;};"
            saveBtn2.OnClientClick = "if (confirm(""Are you sure to save this record?"")) {window.document.forms[0].mode.value=""SAVE"";}else{return false;};"

            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then

        End If

        FL_NAME.CssClass = "REQUIRED"
        AR_NAME.CssClass = "REQUIRED"
        RK_NAME.CssClass = "REQUIRED"

        If flTr.Visible = True Then
            If mFL_NUM.Text.Trim = "" Or FL_NAME.Text.Trim = "" Then
                FL_NUM.Value = ViewState("pre_flcode")
                AR_CODE.Value = ViewState("pre_arcode")
                RK_CODE.Value = ViewState("pre_rkcode")
                BN_CODE.Value = ViewState("pre_bncode")
            End If
        ElseIf arTr.Visible = True Then
            If mAR_CODE.Text.Trim = "" Or AR_NAME.Text.Trim = "" Then
                FL_NUM.Value = ViewState("pre_flcode")
                AR_CODE.Value = ViewState("pre_arcode")
                RK_CODE.Value = ViewState("pre_rkcode")
                BN_CODE.Value = ViewState("pre_bncode")
            End If

        ElseIf rkTr.Visible = True Then
            If mRK_CODE.Text.Trim = "" Or RK_NAME.Text.Trim = "" Then
                FL_NUM.Value = ViewState("pre_flcode")
                AR_CODE.Value = ViewState("pre_arcode")
                RK_CODE.Value = ViewState("pre_rkcode")
                BN_CODE.Value = ViewState("pre_bncode")
            End If

        ElseIf bnTr.Visible = True Then
            If mBN_CODE.Text.Trim = "" Then
                FL_NUM.Value = ViewState("pre_flcode")
                AR_CODE.Value = ViewState("pre_arcode")
                RK_CODE.Value = ViewState("pre_rkcode")
                'BN_CODE.Value = ViewState("pre_bncode")
            End If
        End If

        Select Case mode.Value
            Case "addFL"
                AddFloor()
            Case "addAR"
                AddArea()
            Case "addRK"
                AddRack()
            Case "addBN"
                AddBin()
            Case "SAVE"
                Exit Select
            Case Else
                GenerateAll()
        End Select

    End Sub

    Private Sub fillingUpTables(Optional ByVal flag As String = "")
        Dim srch_sql As String = ""

        If WH_CODE.Value <> "" Then
            If ViewState("first_load") Then
                srch_sql = "select *, FL_NAME as FL_NAME_D, '' as mFlag, '' as newPK, '' as remainYN, '' as hasFile, '' as removeFile from WMS_WH_FL WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
                srch_sql += " order by FL_NAME"
                flDt = New DataTable
                flDt = gDB.getDataTable(srch_sql)

                flDt.Constraints.Add("FL_NUM", flDt.Columns("FL_NUM"), True)

                ViewState("first_load") = False
                ViewState("flDt") = flDt
            End If
        End If

        If flag <> "SAVE" And flag <> "ADD" Then
            srch_sql = "select *, AR_NAME as AR_NAME_D, '' as mFlag, '' as newPK, '' as remainYN, '' as hasFile, '' as removeFile from WMS_WH_AREA WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' and FL_NUM = '" & gU.decodeNullOrEmpty(gU.dbEncode(FL_NUM.Value), "NULL") & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
            srch_sql += " order by AR_NAME"
            arDt = New DataTable
            arDt = gDB.getDataTable(srch_sql)
            ViewState("arDt") = arDt

            srch_sql = "select *, RK_NAME as RK_NAME_D, '' as mFlag, '' as newPK, '' as remainYN from WMS_WH_RACK WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' and FL_NUM ='" & gU.decodeNullOrEmpty(gU.dbEncode(FL_NUM.Value), "NULL") & "' AND AR_CODE = '" & gU.dbEncode(AR_CODE.Value) & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
            srch_sql += " order by RK_NAME"
            rkDt = New DataTable
            rkDt = gDB.getDataTable(srch_sql)
            ViewState("rkDt") = rkDt
        End If

        'If ViewState("loadbn") Then
        If flag <> "SAVE" And flag <> "ADD" Then
            Dim dt As New DataTable
            bnDt = New DataTable

            srch_sql = "select IsNULL(RK_X,0)RK_X, IsNUll(RK_Y,0)RK_Y from WMS_WH_RACK WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' and FL_NUM = '" & gU.decodeNullOrEmpty(gU.dbEncode(FL_NUM.Value), "NULL") & "' AND AR_CODE = '" & gU.dbEncode(AR_CODE.Value) & "' AND RK_CODE = '" & gU.dbEncode(RK_CODE.Value) & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
            dt = gDB.getDataTable(srch_sql)

            If dt.Rows.Count > 0 Then
                For i As Integer = 1 To gU.decodeNull(dt.Rows(0).Item(0), 0)
                    bnDt.Columns.Add("COLUMN" & i)
                Next
                For i As Integer = 1 To gU.decodeNull(dt.Rows(0).Item(1), 0)
                    bnDt.Rows.Add()
                Next
            End If

            srch_sql = "select *,BN_CODE as BN_CODE_D, '' as mFlag, '' as newPK, '' as remainYN from WMS_WH_BIN WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' and FL_NUM ='" & gU.decodeNullOrEmpty(gU.dbEncode(FL_NUM.Value), "NULL") & "' AND AR_CODE = '" & gU.dbEncode(AR_CODE.Value) & "' AND RK_CODE = '" & gU.dbEncode(RK_CODE.Value) & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
            srch_sql += " order by BN_CODE"
            dt = New DataTable
            dt = gDB.getDataTable(srch_sql)

            For i As Integer = 0 To dt.Rows.Count - 1
                If dt.Rows(i).Item("BN_Y") Is DBNull.Value Then
                    dt.Rows(i).Item("BN_Y") = "0"
                End If
                If dt.Rows(i).Item("BN_X") Is DBNull.Value Then
                    dt.Rows(i).Item("BN_X") = "0"
                End If
                If gU.decodeNull(dt.Rows(i).Item("BN_Y"), 0) > 0 And gU.decodeNull(dt.Rows(i).Item("BN_X"), 0) > 0 Then
                    bnDt.Rows(dt.Rows(i).Item("BN_Y") - 1).Item("COLUMN" & dt.Rows(i).Item("BN_X")) = dt.Rows(i).Item("BN_CODE").ToString
                End If

                If Not bnDict.ContainsKey((FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & dt.Rows(i).Item("BN_CODE").ToString)) Then
                    bnDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & dt.Rows(i).Item("BN_CODE").ToString), dt.Rows(i))
                    'Session("bnDict") = bnDict
                End If
            Next
            ViewState("bnsaveDt") = dt
            ViewState("loadbn") = False
            ViewState("bnDt") = bnDt
        End If


    End Sub

    Private Sub GenerateAll(Optional ByVal reloadFlag As String = "")
        If ViewState("pre_flcode") <> FL_NUM.Value Or _
            ViewState("pre_arcode") <> AR_CODE.Value Or _
            ViewState("pre_rkcode") <> RK_CODE.Value Then
            ViewState("loadbn") = True
        Else
            bnDt = ViewState("bnDt")
        End If

        Dim colCount As Integer = 0

        Dim table As New Table

        Dim nbutton As Button
        Dim type As Integer = 0

        flTr.Visible = False
        arTr.Visible = False
        rkTr.Visible = False
        bnTr.Visible = False

        table.BorderWidth = 0
        table.CellPadding = 1
        table.CellSpacing = 1

        table.Style.Add("width", "60%")
        table.HorizontalAlign = HorizontalAlign.Center

        table.ID = "headerTbl"

        locDiv.Controls.Add(table)

        Dim headerrow As New TableRow()
        Dim headercell As New TableCell()
        headercell.ColumnSpan = 3

        Dim tit_table As New Table
        tit_table.BorderWidth = 0
        tit_table.CellPadding = 0
        tit_table.CellSpacing = 0
        tit_table.Style.Add("align", "center")
        tit_table.Style.Add("width", "100%")

        Dim tit_row As New TableRow()
        Dim tit_cell1 As New TableCell()

        tit_cell1.CssClass = "TITLE"
        tit_cell1.Style.Add("width", "100%")

        Dim tit_Label As New Label()
        tit_Label.ID = "lblTitle_N"

        If Session("gLang") = "E" Then
            tit_Label.Text = "Location Definition"
        Else
            tit_Label.Text = "Location Definition"
        End If

        tit_Label.Font.Bold = True
        tit_cell1.Controls.Add(tit_Label)

        tit_row.Cells.Add(tit_cell1)
        tit_table.Rows.Add(tit_row)

        headercell.Controls.Add(tit_table)

        headerrow.Cells.Add(headercell)
        table.Rows.Add(headerrow)

        REM Generate Warehouse
        Dim whRow As New TableRow()

        Dim tclblWH_CODE As New TableCell()
        tclblWH_CODE.CssClass = "LabelTD"
        tclblWH_CODE.Style.Add("width", "20%")

        Dim WH_CODElabel As New Label
        WH_CODElabel.Text = "WareHouse"

        tclblWH_CODE.Controls.Add(WH_CODElabel)
        whRow.Cells.Add(tclblWH_CODE)

        Dim tcWH_CODE As New TableCell()
        tcWH_CODE.Style.Add("width", "70%")
        tcWH_CODE.ColumnSpan = 2

        Dim srch_sql As String = ""
        srch_sql = "select WH_CODE, WH_NAME from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"

        If wh.Value <> "" Then
            srch_sql = srch_sql & " AND WH_CODE = '" & gU.dbEncode(wh.Value) & "'"
        End If

        srch_sql += " order by WH_NAME"

        Dim whdt As DataTable = gDB.getDataTable(srch_sql)

        For i As Integer = 0 To whdt.Rows.Count - 1
            nbutton = New Button
            nbutton.Width = 125
            nbutton.CssClass = "loc_button"
            nbutton.ID = "WH_CODE_" & whdt.Rows(i).Item(0).ToString
            nbutton.Enabled = False

            If WH_CODE.Value = whdt.Rows(i).Item(0).ToString Then
                'nbutton.Text = whdt.Rows(i).Item(1).ToString & " * "
                nbutton.Text = whdt.Rows(i).Item(1).ToString
                nbutton.CssClass = "selected_button"
                nbutton.Enabled = True
            Else
                nbutton.Text = whdt.Rows(i).Item(1).ToString
            End If

            nbutton.OnClientClick = "window.document.forms[0].WH_CODE.value=""" & HttpUtility.HtmlEncode(whdt.Rows(i).Item(0).ToString) & """;window.document.forms[0].FL_NUM.value="""";window.document.forms[0].AR_CODE.value="""";window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"
            'AddHandler nbutton.Click, AddressOf WHCODEClick

            tcWH_CODE.Controls.Add(nbutton)

            colCount += 1

            If colCount = 3 Then
                Dim lit As New Literal
                lit.Text = "<br/>"
                tcWH_CODE.Controls.Add(lit)
                colCount = 0
            End If
        Next

        whRow.Cells.Add(tcWH_CODE)

        fillingUpTables(reloadFlag)

        REM Generate FL
        colCount = 0

        Dim flRow As New TableRow()

        Dim tclbltrFL_NUM As New TableCell()
        tclbltrFL_NUM.CssClass = "LabelTD"
        tclbltrFL_NUM.Style.Add("width", "20%")

        Dim FL_NUMlabel As New Label
        FL_NUMlabel.Text = "Floor"

        tclbltrFL_NUM.Controls.Add(FL_NUMlabel)
        flRow.Cells.Add(tclbltrFL_NUM)

        Dim tcFL_NUM As New TableCell()
        tcFL_NUM.Style.Add("width", "70%")

        For i As Integer = 0 To flDt.Rows.Count - 1
            nbutton = New Button
            nbutton.Width = 75
            nbutton.CssClass = "loc_button"
            nbutton.ID = "FL_NUM_" & flDt.Rows(i).Item("FL_NUM").ToString

            If FL_NUM.Value = flDt.Rows(i).Item("FL_NUM").ToString Then
                'nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString & " * "
                nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString
                nbutton.CssClass = "selected_button"
                type = 1
            Else
                nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString
            End If

            If flDt.Rows(i).Item("mFlag").ToString = "N" Then
                If FL_NUM.Value = flDt.Rows(i).Item("FL_NUM").ToString Then
                    'StoreData(1, flDt.Rows(i))
                ElseIf ViewState("pre_flcode").ToString = flDt.Rows(i).Item("FL_NUM").ToString Then
                    Dim ran As New Random
                    Dim ranNum As Integer
                    Dim dCount As Integer = 1

                    Do While dCount > 0
                        ranNum = CInt(Now.Millisecond) + CInt(Now.Minute) + ran.Next(1, 9999)
                        dCount = flDt.Compute("Count(FL_NUM)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = '" & ranNum & "'")
                    Loop

                    If StoreData(1, flDt.Rows(i), ranNum) = False Then
                        FL_NUM.Value = flDt.Rows(i).Item("FL_NUM").ToString
                        fillingUpTables(reloadFlag)

                        If flDt.Rows(i).Item("remainYN").ToString = "Y" Then
                            type = 1
                        End If

                        For x As Integer = 0 To tcFL_NUM.Controls.Count - 1
                            Dim tempButton = New Button
                            tempButton = DirectCast(tcFL_NUM.Controls(x), Button)

                            'If tempButton.Text.Contains("*") Then tempButton.Text = tempButton.Text.Replace(" *", "")

                            'nbutton.Text = nbutton.Text.Replace(" *", "") & " *"
                            If tempButton.Text = AR_CODE.Value Then tempButton.CssClass = "selected_button"
                        Next
                    Else
                        If flDt.Rows(i).Item("FL_NAME_D").ToString <> "" Then
                            If FL_NUM.Value = flDt.Rows(i).Item("FL_NUM").ToString Then
                                'nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString & " * "
                                nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString
                                nbutton.CssClass = "selected_button"
                            Else
                                nbutton.Text = flDt.Rows(i).Item("FL_NAME_D").ToString
                            End If

                        Else
                            If FL_NUM.Value = flDt.Rows(i).Item("FL_NUM").ToString Then
                                'nbutton.Text = "NEW" & " * "
                                nbutton.Text = "NEW"
                                nbutton.CssClass = "selected_button"
                            Else
                                nbutton.Text = "NEW"
                            End If
                        End If
                    End If
                End If
            End If

            nbutton.OnClientClick = "window.document.forms[0].FL_NUM.value=""" & HttpUtility.HtmlEncode(flDt.Rows(i).Item("FL_NUM").ToString) & """;window.document.forms[0].AR_CODE.value="""";window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"

            tcFL_NUM.Controls.Add(nbutton)

            colCount += 1

            If colCount = 5 Then
                Dim lit As New Literal
                lit.Text = "<br/>"
                tcFL_NUM.Controls.Add(lit)
                colCount = 0
            End If
        Next

        flRow.Cells.Add(tcFL_NUM)

        Dim tcFL_NUM_NEW As New TableCell()
        tcFL_NUM_NEW.Style.Add("width", "10%")

        If WH_CODE.Value <> "" Then
            Dim flNewBtn As New Button
            flNewBtn.Text = "+ Floor"
            flNewBtn.CssClass = "add_loc_button"
            flNewBtn.OnClientClick = "window.document.forms[0].mode.value=""addFL"";"
            tcFL_NUM_NEW.Controls.Add(flNewBtn)
        End If

        flRow.Cells.Add(tcFL_NUM_NEW)

        REM Generate AR
        colCount = 0
        Dim arRow As New TableRow()

        Dim tclbltrAR_CODE As New TableCell()
        tclbltrAR_CODE.CssClass = "LabelTD"
        tclbltrAR_CODE.Style.Add("width", "20%")

        Dim AR_CODElabel As New Label
        AR_CODElabel.Text = "Area"

        tclbltrAR_CODE.Controls.Add(AR_CODElabel)
        arRow.Cells.Add(tclbltrAR_CODE)

        Dim tcAR_CODE As New TableCell()
        tcAR_CODE.Style.Add("width", "70%")

        Dim oldCode As String = ""
        Dim codeArray As New ArrayList

        If arDict.Count = 0 And arDt.Rows.Count > 0 Then
            For i As Integer = 0 To arDt.Rows.Count - 1
                nbutton = New Button
                nbutton.Width = 75
                nbutton.CssClass = "loc_button"
                nbutton.ID = "AR_CODE_" & arDt.Rows(i).Item("AR_CODE").ToString

                If Not arDict.ContainsKey((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString)) Then
                    arDict.Add((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString), arDt.Rows(i))
                    'Session("arDict") = arDict
                End If

                If AR_CODE.Value = arDt.Rows(i).Item("AR_CODE").ToString Then
                    'nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString & " * "
                    nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString
                    nbutton.CssClass = "selected_button"
                    type = 2
                Else
                    nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString
                End If

                nbutton.OnClientClick = "window.document.forms[0].AR_CODE.value=""" & HttpUtility.HtmlEncode(arDt.Rows(i).Item("AR_CODE").ToString) & """;window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"

                tcAR_CODE.Controls.Add(nbutton)

                colCount += 1

                If colCount = 5 Then
                    Dim lit As New Literal
                    lit.Text = "<br/>"
                    tcAR_CODE.Controls.Add(lit)
                    colCount = 0
                End If
            Next
        ElseIf arDict.Count > 0 Then
            If arDt.Rows.Count > 0 Then
                For i As Integer = 0 To arDt.Rows.Count - 1
                    nbutton = New Button
                    nbutton.Width = 75
                    nbutton.CssClass = "loc_button"
                    nbutton.ID = "AR_CODE_" & arDt.Rows(i).Item("AR_CODE").ToString

                    If Not arDict.ContainsKey((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString)) Then
                        arDict.Add((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString), arDt.Rows(i))
                        'Session("arDict") = arDict
                    End If

                    If AR_CODE.Value = arDt.Rows(i).Item("AR_CODE").ToString Then
                        'nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString & " * "
                        nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString
                        nbutton.CssClass = "selected_button"
                        type = 2
                    Else
                        nbutton.Text = arDt.Rows(i).Item("AR_NAME_D").ToString
                    End If

                    nbutton.OnClientClick = "window.document.forms[0].AR_CODE.value=""" & HttpUtility.HtmlEncode(arDt.Rows(i).Item("AR_CODE").ToString) & """;window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"

                    tcAR_CODE.Controls.Add(nbutton)

                    colCount += 1

                    If colCount = 5 Then
                        Dim lit As New Literal
                        lit.Text = "<br/>"
                        tcAR_CODE.Controls.Add(lit)
                        colCount = 0
                    End If
                Next

                Dim tempRow As DataRow
                For Each arkeyPair As KeyValuePair(Of String, DataRow) In arDict
                    Dim tempKey As String = arkeyPair.Key
                    Dim keysStr As String() = Split(tempKey, "||")

                    If FL_NUM.Value = "" Then
                        Exit For
                    End If

                    If keysStr(0) = FL_NUM.Value Then
                        Dim existArCnt As Integer = 0

                        existArCnt = arDt.Compute("Count(AR_CODE)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = '" & keysStr(0) & "' AND AR_CODE = '" & keysStr(1) & "'")

                        If existArCnt = 0 Then

                            tempRow = arkeyPair.Value

                            nbutton = New Button
                            nbutton.Width = 75
                            nbutton.CssClass = "loc_button"
                            nbutton.ID = "AR_CODE_" & tempRow.Item("AR_CODE").ToString

                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                'nbutton.Text = tempRow.Item("AR_NAME_D").ToString & " * "
                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                nbutton.CssClass = "selected_button"
                                type = 2
                            Else
                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                            End If

                            If tempRow.Item("mFlag").ToString = "N" Then
                                If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                    'StoreData(2, tempRow)
                                ElseIf ViewState("pre_arcode").ToString = tempRow.Item("AR_CODE").ToString Then
                                    If StoreData(2, tempRow, , oldCode) = False Then
                                        AR_CODE.Value = tempRow.Item("AR_CODE").ToString
                                        fillingUpTables(reloadFlag)

                                        If tempRow.Item("remainYN").ToString = "Y" Then
                                            type = 2
                                        End If

                                        For x As Integer = 0 To tcAR_CODE.Controls.Count - 1
                                            Dim tempButton = New Button
                                            tempButton = DirectCast(tcAR_CODE.Controls(x), Button)

                                            'If tempButton.Text.Contains("*") Then tempButton.Text = tempButton.Text.Replace(" *", "")

                                            'nbutton.Text = nbutton.Text.Replace(" *", "") & " *"
                                            If tempButton.Text = AR_CODE.Value Then tempButton.CssClass = "selected_button"
                                        Next
                                    Else
                                        If oldCode <> "" Then codeArray.Add(oldCode)

                                        If tempRow.Item("AR_NAME_D").ToString <> "" Then
                                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                                'nbutton.Text = tempRow.Item("AR_NAME_D").ToString & " * "
                                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                                nbutton.CssClass = "selected_button"
                                            Else
                                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                            End If
                                        Else
                                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                                'nbutton.Text = "NEW" & " * "
                                                nbutton.Text = "NEW"
                                                nbutton.CssClass = "selected_button"
                                            Else
                                                nbutton.Text = "NEW"
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            nbutton.OnClientClick = "window.document.forms[0].AR_CODE.value=""" & HttpUtility.HtmlEncode(tempRow.Item("AR_CODE").ToString) & """;window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"

                            tcAR_CODE.Controls.Add(nbutton)

                            colCount += 1

                            If colCount = 5 Then
                                Dim lit As New Literal
                                lit.Text = "<br/>"
                                tcAR_CODE.Controls.Add(lit)
                                colCount = 0
                            End If
                        End If
                    End If
                Next
            Else
                Dim tempRow As DataRow
                For Each arkeyPair As KeyValuePair(Of String, DataRow) In arDict
                    Dim tempKey As String = arkeyPair.Key
                    Dim keysStr As String() = Split(tempKey, "||")

                    If FL_NUM.Value = "" Then
                        Exit For
                    End If

                    If keysStr(0) = FL_NUM.Value Then
                        Dim existArCnt As Integer = 0

                        existArCnt = arDt.Compute("Count(AR_CODE)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = " & keysStr(0) & " AND AR_CODE = '" & keysStr(1) & "'")

                        If existArCnt = 0 Then
                            tempRow = arkeyPair.Value

                            nbutton = New Button
                            nbutton.Width = 75
                            nbutton.CssClass = "loc_button"
                            nbutton.ID = "AR_CODE_" & tempRow.Item("AR_CODE").ToString

                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                'nbutton.Text = tempRow.Item("AR_NAME_D").ToString & " * "
                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                nbutton.CssClass = "selected_button"
                                type = 2
                            Else
                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                            End If

                            If tempRow.Item("mFlag").ToString = "N" Then

                                If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                    'StoreData(2, tempRow)
                                ElseIf ViewState("pre_arcode").ToString = tempRow.Item("AR_CODE").ToString Then
                                    If StoreData(2, tempRow, , oldCode) = False Then
                                        AR_CODE.Value = tempRow.Item("AR_CODE").ToString
                                        fillingUpTables(reloadFlag)

                                        If tempRow.Item("remainYN").ToString = "Y" Then
                                            type = 2
                                        End If

                                        For x As Integer = 0 To tcAR_CODE.Controls.Count - 1
                                            Dim tempButton = New Button
                                            tempButton = DirectCast(tcAR_CODE.Controls(x), Button)

                                            'If tempButton.Text.Contains("*") Then tempButton.Text = tempButton.Text.Replace(" *", "")

                                            'nbutton.Text = nbutton.Text.Replace(" *", "") & " *"

                                            If tempButton.Text = AR_CODE.Value Then tempButton.CssClass = "selected_button"
                                        Next
                                    Else
                                        If oldCode <> "" Then codeArray.Add(oldCode)

                                        If tempRow.Item("AR_NAME_D").ToString <> "" Then
                                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                                'nbutton.Text = tempRow.Item("AR_NAME_D").ToString & " * "
                                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                                nbutton.CssClass = "selected_button"
                                            Else
                                                nbutton.Text = tempRow.Item("AR_NAME_D").ToString
                                            End If
                                        Else
                                            If AR_CODE.Value = tempRow.Item("AR_CODE").ToString Then
                                                'nbutton.Text = nbutton.Text = "NEW" & " * "
                                                nbutton.Text = nbutton.Text = "NEW"
                                                nbutton.CssClass = "selected_button"
                                            Else
                                                nbutton.Text = nbutton.Text = "NEW"
                                            End If

                                        End If
                                    End If
                                End If
                            End If

                            nbutton.OnClientClick = "window.document.forms[0].AR_CODE.value=""" & HttpUtility.HtmlEncode(tempRow.Item("AR_CODE").ToString) & """;window.document.forms[0].RK_CODE.value="""";window.document.forms[0].BN_CODE.value="""";"

                            tcAR_CODE.Controls.Add(nbutton)

                            colCount += 1

                            If colCount = 5 Then
                                Dim lit As New Literal
                                lit.Text = "<br/>"
                                tcAR_CODE.Controls.Add(lit)
                                colCount = 0
                            End If
                        End If
                    End If
                Next
            End If
        End If

        If codeArray.Count <> 0 Then
            For x As Integer = 0 To codeArray.Count - 1
                If arDict.ContainsKey(codeArray(x)) Then
                    Dim codeRow As DataRow = arDict.Item(codeArray(x))

                    If codeRow IsNot Nothing Then
                        Dim dictKey As String = codeRow.Item("FL_NUM").ToString & "||" & codeRow.Item("AR_CODE").ToString
                        arDict.Remove(codeArray(x))
                        arDict.Add(dictKey, codeRow)
                    End If
                End If
            Next
        End If

        arRow.Cells.Add(tcAR_CODE)

        Dim tcAR_CODE_NEW As New TableCell()
        tcAR_CODE_NEW.Style.Add("width", "10%")

        If WH_CODE.Value <> "" And FL_NUM.Value <> "" Then
            Dim arNewBtn As New Button
            arNewBtn.Text = "+ Area"
            arNewBtn.CssClass = "add_loc_button"
            arNewBtn.OnClientClick = "window.document.forms[0].mode.value=""addAR"";"
            tcAR_CODE_NEW.Controls.Add(arNewBtn)
        End If

        arRow.Cells.Add(tcAR_CODE_NEW)

        REM Generate Rack
        colCount = 0
        Dim rackRow As New TableRow()

        Dim tclbltrRK_CODE As New TableCell()
        tclbltrRK_CODE.CssClass = "LabelTD"
        tclbltrRK_CODE.Style.Add("width", "20%")

        Dim RK_CODElabel As New Label
        RK_CODElabel.Text = "Rack"

        tclbltrRK_CODE.Controls.Add(RK_CODElabel)
        rackRow.Cells.Add(tclbltrRK_CODE)

        Dim tcRK_CODE As New TableCell()
        tcRK_CODE.Style.Add("width", "70%")

        codeArray = New ArrayList

        If rkDict.Count = 0 And rkDt.Rows.Count > 0 Then
            For i As Integer = 0 To rkDt.Rows.Count - 1
                nbutton = New Button
                nbutton.Width = 75
                nbutton.CssClass = "loc_button"
                nbutton.ID = "RK_CODE_" & rkDt.Rows(i).Item("RK_CODE").ToString

                If Not rkDict.ContainsKey((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString)) Then
                    rkDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString), rkDt.Rows(i))
                    'Session("rkDict") = rkDict
                End If

                If RK_CODE.Value = rkDt.Rows(i).Item("RK_CODE").ToString Then
                    'nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString & " * "
                    nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString
                    nbutton.CssClass = "selected_button"
                    type = 3
                Else
                    nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString
                End If

                nbutton.OnClientClick = "window.document.forms[0].RK_CODE.value=""" & HttpUtility.HtmlEncode(rkDt.Rows(i).Item("RK_CODE").ToString) & """;window.document.forms[0].BN_CODE.value="""";"

                tcRK_CODE.Controls.Add(nbutton)

                colCount += 1

                If colCount = 5 Then
                    Dim lit As New Literal
                    lit.Text = "<br/>"
                    tcRK_CODE.Controls.Add(lit)
                    colCount = 0
                End If
            Next
        ElseIf rkDict.Count > 0 Then
            If rkDt.Rows.Count > 0 Then
                For i As Integer = 0 To rkDt.Rows.Count - 1
                    nbutton = New Button
                    nbutton.Width = 75
                    nbutton.CssClass = "loc_button"
                    nbutton.ID = "RK_CODE_" & rkDt.Rows(i).Item("RK_CODE").ToString

                    If Not rkDict.ContainsKey((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString)) Then
                        rkDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString), rkDt.Rows(i))
                        'Session("rkDict") = rkDict
                    End If

                    If RK_CODE.Value = rkDt.Rows(i).Item("RK_CODE").ToString Then
                        'nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString & " * "
                        nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString
                        nbutton.CssClass = "selected_button"
                        type = 3
                    Else
                        nbutton.Text = rkDt.Rows(i).Item("RK_NAME_D").ToString
                    End If

                    nbutton.OnClientClick = "window.document.forms[0].RK_CODE.value=""" & HttpUtility.HtmlEncode(rkDt.Rows(i).Item("RK_CODE").ToString) & """;window.document.forms[0].BN_CODE.value="""";"

                    tcRK_CODE.Controls.Add(nbutton)

                    colCount += 1

                    If colCount = 5 Then
                        Dim lit As New Literal
                        lit.Text = "<br/>"
                        tcRK_CODE.Controls.Add(lit)
                        colCount = 0
                    End If
                Next

                Dim tempRow As DataRow
                For Each rkkeyPair As KeyValuePair(Of String, DataRow) In rkDict
                    Dim tempKey As String = rkkeyPair.Key
                    Dim keysStr As String() = Split(tempKey, "||")

                    If FL_NUM.Value = "" Or AR_CODE.Value = "" Then
                        Exit For
                    End If

                    If keysStr(0) = FL_NUM.Value And keysStr(1) = AR_CODE.Value Then
                        Dim existArCnt As Integer = 0

                        existArCnt = rkDt.Compute("Count(RK_CODE)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = '" & keysStr(0) & "' AND AR_CODE = '" & keysStr(1) & "' AND RK_CODE = '" & keysStr(2) & "'")
                        If existArCnt = 0 Then

                            tempRow = rkkeyPair.Value

                            nbutton = New Button
                            nbutton.Width = 75
                            nbutton.CssClass = "loc_button"
                            nbutton.ID = "RK_CODE_" & tempRow.Item("RK_CODE").ToString

                            If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                'nbutton.Text = tempRow.Item("RK_NAME_D").ToString & " * "
                                nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                nbutton.CssClass = "selected_button"
                                type = 3
                            Else
                                nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                            End If

                            If tempRow.Item("mFlag").ToString = "N" Then
                                If RK_CODE.Value <> tempRow.Item("RK_CODE").ToString Then
                                    If ViewState("pre_rkcode").ToString = tempRow.Item("RK_CODE").ToString Then
                                        If StoreData(3, tempRow, , oldCode) = False Then
                                            RK_CODE.Value = tempRow.Item("RK_CODE").ToString
                                            fillingUpTables(reloadFlag)

                                            If tempRow.Item("remainYN").ToString = "Y" Then
                                                type = 3
                                            End If

                                            For x As Integer = 0 To tcRK_CODE.Controls.Count - 1
                                                Dim tempButton = New Button
                                                tempButton = DirectCast(tcRK_CODE.Controls(x), Button)

                                                'If tempButton.Text.Contains("*") Then tempButton.Text = tempButton.Text.Replace(" *", "")

                                                'nbutton.Text = nbutton.Text.Replace(" *", "") & " *"
                                                If tempButton.Text = AR_CODE.Value Then tempButton.CssClass = "selected_button"
                                            Next
                                        Else
                                            If oldCode <> "" Then codeArray.Add(oldCode)

                                            If tempRow.Item("RK_NAME_D").ToString <> "" Then
                                                If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                                    'nbutton.Text = tempRow.Item("RK_NAME_D").ToString & " * "
                                                    nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                                    nbutton.CssClass = "selected_button"
                                                Else
                                                    nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                                End If
                                            Else
                                                If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                                    'nbutton.Text = "NEW" & " * "
                                                    nbutton.Text = "NEW"
                                                    nbutton.CssClass = "selected_button"
                                                Else
                                                    nbutton.Text = "NEW"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            nbutton.OnClientClick = "window.document.forms[0].RK_CODE.value=""" & HttpUtility.HtmlEncode(tempRow.Item("RK_CODE").ToString) & """;window.document.forms[0].BN_CODE.value="""";"

                            tcRK_CODE.Controls.Add(nbutton)

                            colCount += 1

                            If colCount = 5 Then
                                Dim lit As New Literal
                                lit.Text = "<br/>"
                                tcRK_CODE.Controls.Add(lit)
                                colCount = 0
                            End If
                        End If
                    End If
                Next
            Else
                Dim tempRow As DataRow
                For Each rkkeyPair As KeyValuePair(Of String, DataRow) In rkDict
                    Dim tempKey As String = rkkeyPair.Key
                    Dim keysStr As String() = Split(tempKey, "||")

                    If FL_NUM.Value = "" Or AR_CODE.Value = "" Then
                        Exit For
                    End If

                    If keysStr(0) = FL_NUM.Value And keysStr(1) = AR_CODE.Value Then
                        Dim existArCnt As Integer = 0

                        existArCnt = rkDt.Compute("Count(RK_CODE)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = " & keysStr(0) & _
                                                                    " AND AR_CODE = '" & keysStr(1) & "' AND RK_CODE = '" & keysStr(2) & "'")
                        If existArCnt = 0 Then

                            tempRow = rkkeyPair.Value

                            nbutton = New Button
                            nbutton.Width = 75
                            nbutton.CssClass = "loc_button"
                            nbutton.ID = "RK_CODE_" & tempRow.Item("RK_CODE").ToString

                            If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                'nbutton.Text = tempRow.Item("RK_NAME_D").ToString & " * "
                                nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                nbutton.CssClass = "selected_button"
                                type = 3
                            Else
                                nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                            End If

                            If tempRow.Item("mFlag").ToString = "N" Then
                                If RK_CODE.Value <> tempRow.Item("RK_CODE").ToString Then
                                    If ViewState("pre_rkcode").ToString = tempRow.Item("RK_CODE").ToString Then
                                        If StoreData(3, tempRow, , oldCode) = False Then
                                            RK_CODE.Value = tempRow.Item("RK_CODE").ToString
                                            fillingUpTables(reloadFlag)

                                            If tempRow.Item("remainYN").ToString = "Y" Then
                                                type = 3
                                            End If

                                            For x As Integer = 0 To tcRK_CODE.Controls.Count - 1
                                                Dim tempButton = New Button
                                                tempButton = DirectCast(tcRK_CODE.Controls(x), Button)

                                                'If tempButton.Text.Contains("*") Then tempButton.Text = tempButton.Text.Replace(" *", "")

                                                'nbutton.Text = nbutton.Text.Replace(" *", "") & " *"
                                                If tempButton.Text = AR_CODE.Value Then tempButton.CssClass = "selected_button"
                                            Next
                                        Else
                                            If oldCode <> "" Then codeArray.Add(oldCode)

                                            If tempRow.Item("RK_NAME_D").ToString <> "" Then
                                                If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                                    'nbutton.Text = tempRow.Item("RK_NAME_D").ToString & " * "
                                                    nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                                    nbutton.CssClass = "selected_button"
                                                Else
                                                    nbutton.Text = tempRow.Item("RK_NAME_D").ToString
                                                End If
                                            Else
                                                If RK_CODE.Value = tempRow.Item("RK_CODE").ToString Then
                                                    'nbutton.Text = "NEW" & " * "
                                                    nbutton.Text = "NEW"
                                                    nbutton.CssClass = "selected_button"
                                                Else
                                                    nbutton.Text = "NEW"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            nbutton.OnClientClick = "window.document.forms[0].RK_CODE.value=""" & HttpUtility.HtmlEncode(tempRow.Item("RK_CODE").ToString) & """;window.document.forms[0].BN_CODE.value="""";"

                            tcRK_CODE.Controls.Add(nbutton)

                            colCount += 1

                            If colCount = 5 Then
                                Dim lit As New Literal
                                lit.Text = "<br/>"
                                tcRK_CODE.Controls.Add(lit)
                                colCount = 0
                            End If
                        End If
                    End If
                Next
            End If
        End If

        If codeArray.Count <> 0 Then
            For x As Integer = 0 To codeArray.Count - 1
                If rkDict.ContainsKey(codeArray(x)) Then
                    Dim codeRow As DataRow = rkDict.Item(codeArray(x))

                    If codeRow IsNot Nothing Then
                        Dim dictKey As String = codeRow.Item("FL_NUM").ToString & "||" & codeRow.Item("AR_CODE").ToString & "||" & codeRow.Item("RK_CODE").ToString
                        rkDict.Remove(codeArray(x))
                        rkDict.Add(dictKey, codeRow)
                    End If
                End If
            Next
        End If

        rackRow.Cells.Add(tcRK_CODE)

        Dim tcRK_CODE_NEW As New TableCell()
        tcRK_CODE_NEW.Style.Add("width", "10%")

        If WH_CODE.Value <> "" And FL_NUM.Value <> "" And AR_CODE.Value <> "" Then
            Dim rkNewBtn As New Button
            rkNewBtn.Text = "+ Rack"
            rkNewBtn.CssClass = "add_loc_button"
            rkNewBtn.OnClientClick = "window.document.forms[0].mode.value=""addRK"";"
            tcRK_CODE_NEW.Controls.Add(rkNewBtn)
        End If

        rackRow.Cells.Add(tcRK_CODE_NEW)

        REM Generate BIN
        Dim binRow As New TableRow

        Dim tclbltrBN_CODE As New TableCell()
        tclbltrBN_CODE.CssClass = "LabelTD"
        tclbltrBN_CODE.Style.Add("width", "20%")

        Dim BN_CODElabel As New Label
        BN_CODElabel.Text = "BIN"

        tclbltrBN_CODE.Controls.Add(BN_CODElabel)
        binRow.Cells.Add(tclbltrBN_CODE)

        Dim tcBN_CODE As New TableCell()
        tcBN_CODE.Style.Add("width", "80%")
        tcBN_CODE.ColumnSpan = 2
        REM Create dt for BIN

        Dim newTable As New Table

        Dim preBNCode As String = ViewState("pre_bncode")

        Dim rkC As String = ""

        If rkTr.Visible = True Then
            rkC = mRK_CODE.Text
            RK_CODE.Value = rkC
        Else
            rkC = RK_CODE.Value
        End If

        If bnDict.Keys.Contains(FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkC & "||" & preBNCode) Then
            Dim nRow As DataRow

            nRow = bnDict.Item(FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkC & "||" & preBNCode)

            If nRow.Item("mFlag").ToString = "N" And nRow.Item("newPK").ToString = "Y" And mBN_CODE.Text = "" Then
                bnDt.Rows(CInt(nRow.Item("BN_Y")) - 1).Item(CInt(nRow.Item("BN_X")) - 1) = ""

                bnDict.Remove(FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkC & "||" & preBNCode)
            End If

        End If

        bnDt.AcceptChanges()

        ViewState("bnDt") = bnDt

        fillingUpTables(reloadFlag)

        For i As Integer = 0 To bnDt.Rows.Count - 1
            Dim newRow As New TableRow
            For i1 As Integer = 0 To bnDt.Columns.Count - 1
                Dim newCell As New TableCell
                nbutton = New Button
                nbutton.Width = 75
                nbutton.CssClass = "loc_button"

                If bnDt.Rows(i).Item(i1).ToString = "" Then
                    nbutton.Text = "+"
                    nbutton.CssClass = "add_loc_button"
                    'nbutton.Enabled = False
                    nbutton.ID = "BN_CODE_E_" & i.ToString & "_" & i1.ToString

                    nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & i.ToString & "_E_" & i1.ToString & """;window.document.forms[0].mode.value=""addBN"";")
                Else
                    Dim gotCode As Boolean = True

                    If bnDict.ContainsKey(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnDt.Rows(i).Item(i1).ToString) Then
                        Dim bnRow As DataRow = bnDict.Item(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnDt.Rows(i).Item(i1).ToString)
                        If bnRow.Item("mFlag").ToString = "N" Then
                            If BN_CODE.Value <> bnRow.Item("BN_CODE").ToString Then
                                If BN_CODE.Value = "" Then
                                    nbutton.Text = "+"
                                    nbutton.CssClass = "add_loc_button"
                                    'nbutton.Enabled = False
                                    nbutton.ID = "BN_CODE_E_" & i.ToString & "_" & i1.ToString

                                    nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & i.ToString & "_E_" & i1.ToString & """;window.document.forms[0].mode.value=""addBN"";")

                                    gotCode = False
                                Else
                                    If ViewState("pre_bncode").ToString = bnRow.Item("BN_CODE").ToString Then
                                        If StoreData(4, bnRow, , oldCode) = False Then
                                            nbutton.Text = "+"
                                            nbutton.CssClass = "add_loc_button"

                                            nbutton.ID = "BN_CODE_E_" & i.ToString & "_" & i1.ToString

                                            nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & i.ToString & "_E_" & i1.ToString & """;window.document.forms[0].mode.value=""addBN"";")

                                            bnRow.Item("BN_CODE_D") = ""

                                            gotCode = False
                                        Else

                                            If Not bnDict.ContainsKey(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnRow.Item("BN_CODE").ToString) Then
                                                bnDict.Add(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnRow.Item("BN_CODE").ToString, bnRow)
                                                bnDt.Rows(i).Item(i1) = bnRow.Item("BN_CODE").ToString
                                            End If

                                            If bnDict.ContainsKey(oldCode) Then
                                                bnDict.Remove(oldCode)
                                            End If
                                        End If
                                    ElseIf mBN_CODE.Text = "" And bnRow.Item("newPK").ToString = "Y" Then
                                        nbutton.Text = "+"
                                        nbutton.CssClass = "add_loc_button"

                                        nbutton.ID = "BN_CODE_E_" & i.ToString & "_" & i1.ToString

                                        nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & i.ToString & "_E_" & i1.ToString & """;window.document.forms[0].mode.value=""addBN"";")

                                        bnRow.Item("BN_CODE_D") = ""

                                        gotCode = False
                                    Else
                                        'gotCode = False
                                    End If
                                End If

                            End If

                        End If

                        If gotCode Then
                            nbutton.ID = "BN_CODE_" & bnDt.Rows(i).Item(i1).ToString

                            If BN_CODE.Value = bnDt.Rows(i).Item(i1).ToString Then
                                'nbutton.Text = bnRow.Item("BN_CODE_D").ToString & " * "
                                nbutton.Text = bnRow.Item("BN_CODE_D").ToString
                                nbutton.CssClass = "selected_button"
                                type = 4

                                nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & HttpUtility.HtmlEncode(bnDt.Rows(i).Item(i1).ToString.Trim) & """;")
                            Else
                                If bnRow.Item("BN_CODE_D").ToString = "" Then
                                    nbutton.Text = "+"
                                    nbutton.CssClass = "add_loc_button"

                                    nbutton.ID = "BN_CODE_E_" & i.ToString & "_" & i1.ToString

                                    nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & i.ToString & "_E_" & i1.ToString & """;window.document.forms[0].mode.value=""addBN"";")
                                Else
                                    nbutton.Text = bnRow.Item("BN_CODE_D").ToString

                                    nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & HttpUtility.HtmlEncode(bnDt.Rows(i).Item(i1).ToString.Trim) & """;")
                                End If

                            End If


                        End If
                    End If


                End If

                nbutton.Style.Add("height", "50px")
                nbutton.Style.Add("width", "50px")
                'nbutton.Width = "50"
                'nbutton.Height = "50"
                newCell.Controls.Add(nbutton)
                newRow.Cells.Add(newCell)
            Next
            newTable.Rows.Add(newRow)
        Next
        tcBN_CODE.Controls.Add(newTable)

        'trBN_CODE.Cells.Add(tcBN_CODE)

        binRow.Cells.Add(tcBN_CODE)

        table.Rows.Add(whRow)
        table.Rows.Add(flRow)
        table.Rows.Add(arRow)
        table.Rows.Add(rackRow)
        table.Rows.Add(binRow)

        If ViewState("pre_flcode").ToString <> "" Or _
            ViewState("pre_arcode").ToString <> "" Or _
             ViewState("pre_rkcode").ToString <> "" Or _
             ViewState("pre_bncode").ToString <> "" Then

            StoreData(ViewState("pre_flcode").ToString, ViewState("pre_arcode").ToString, ViewState("pre_rkcode").ToString, ViewState("pre_bncode").ToString)
        End If

        ShowTable(type)

        ViewState("pre_flcode") = FL_NUM.Value
        ViewState("pre_arcode") = AR_CODE.Value
        ViewState("pre_rkcode") = RK_CODE.Value
        ViewState("pre_bncode") = BN_CODE.Value

    End Sub

    Private Function StoreData(ByVal type As Integer, ByRef row As DataRow, Optional ByVal ranNum As String = "", Optional ByRef old_Code As String = "") As Boolean
        Dim tempCodeStr As String = ""
        Dim tblCode As String = ""

        Select Case type
            Case 1
                If row IsNot Nothing Then
                    row.Item("remainYN") = ""
                    ViewState("notAllowAdd") = False

                    If mFL_NUM.Text.Trim <> "" Then
                        If FL_NAME.Text <> "" Then
                            If mFL_NUM.Text.Trim <> row.Item("FL_NUM").ToString Then
                                Dim nDataRows() As DataRow = flDt.Select("FL_NUM = " & mFL_NUM.Text.Trim)

                                If nDataRows.Count = 0 Then
                                    tblCode = row.Item("FL_NUM").ToString
                                    row.Item("FL_NUM") = gU.decodeEmptyCInt(mFL_NUM.Text.Trim, 0)
                                    row.Item("newPK") = ""
                                Else
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Duplicate Floor has found!!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                    End If

                                    row.Item("remainYN") = "Y"
                                    ViewState("notAllowAdd") = True

                                    Return False
                                End If
                            End If
                        Else
                            row.Item("remainYN") = "Y"
                            ViewState("notAllowAdd") = True

                            Return False
                        End If
                    Else
                        'Dim ranNum As New Random
                        'row.Item("FL_NUM") = 1 + CInt(Now.Millisecond) + CInt(Now.Minute) + ranNum.Next(1, 2999)
                        'row.Item("FL_NUM") = CInt(ranNum)
                        'row.Item("newPK") = "Y"

                        'If Session("gLang") = "E" Then
                        '    uiFun.displayMsg(Me, "", "Duplicate Floor Cannot Be Empty!!", Session("gLang"))
                        'Else
                        '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                        'End If

                        row.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True

                        Return False
                    End If

                    If FL_NAME.Text <> "" Then row.Item("FL_NAME_D") = FL_NAME.Text Else row.Item("FL_NAME_D") = "NEW"
                    row.Item("FL_NAME") = FL_NAME.Text
                    row.Item("FL_NAME_CH") = FL_NAME_CH.Text
                    row.Item("FL_GROSS_AREA") = gU.decodeEmptyCdbl(FL_GROSS_AREA.Text, 0)
                    row.Item("FL_GROSS_CBM") = gU.decodeEmptyCdbl(FL_GROSS_CBM.Text, 0)
                    row.Item("FL_MAX_WEIGHT") = gU.decodeEmptyCdbl(FL_MAX_WEIGHT.Text, 0)
                    row.Item("FL_LENGTH") = gU.decodeEmptyCdbl(FL_LENGTH.Text, 0)
                    row.Item("FL_WIDTH") = gU.decodeEmptyCdbl(FL_WIDTH.Text, 0)
                    row.Item("FL_HEIGHT") = gU.decodeEmptyCdbl(FL_HEIGHT.Text, 0)

                    If FL_PICTURE_upload.HasFile Then
                        If Not flPhoto.ContainsKey(row.Item("FL_NUM").ToString) Then
                            flPhoto.Add(row.Item("FL_NUM").ToString, FL_PICTURE_upload.PostedFile)
                            row.Item("hasFile") = "Y"
                            'Session("flPhoto") = flPhoto
                        Else
                            Dim tempHttpPostFile As HttpPostedFile = flPhoto.Item(row.Item("FL_NUM").ToString)

                            If tempHttpPostFile.FileName <> FL_PICTURE_upload.PostedFile.FileName Then
                                flPhoto.Item(row.Item("FL_NUM").ToString) = FL_PICTURE_upload.PostedFile
                                'Session("flPhoto") = flPhoto
                            End If

                            tempHttpPostFile = Nothing
                        End If
                    End If
                End If
            Case 2
                If row IsNot Nothing Then
                    row.Item("remainYN") = ""
                    ViewState("notAllowAdd") = False

                    If mAR_CODE.Text.Trim <> "" Then
                        If AR_NAME.Text <> "" Then
                            If mAR_CODE.Text <> row.Item("AR_CODE").ToString Then
                                Dim dicCode As String = row.Item("FL_NUM").ToString & "||" & mAR_CODE.Text

                                If Not arDict.ContainsKey(dicCode) Then
                                    tblCode = row.Item("FL_NUM").ToString & "||" & row.Item("AR_CODE").ToString
                                    row.Item("AR_CODE") = mAR_CODE.Text
                                    row.Item("newPK") = ""
                                Else
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Duplicate Area has found!!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                    End If

                                    row.Item("remainYN") = "Y"
                                    ViewState("notAllowAdd") = True

                                    Return False
                                End If
                            End If
                        Else
                            row.Item("remainYN") = "Y"
                            ViewState("notAllowAdd") = True

                            Return False
                        End If
                    Else
                        ''Dim ranNum As New Random
                        'row.Item("AR_CODE") = 1 + CInt(Now.Millisecond) + CInt(Now.Minute) + ranNum.Next(1, 2999)
                        'row.Item("newPK") = "Y"

                        'If Session("gLang") = "E" Then
                        '    uiFun.displayMsg(Me, "", "Duplicate Area Code Cannot Be Empty!!", Session("gLang"))
                        'Else
                        '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                        'End If

                        row.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True

                        Return False
                    End If

                    'row.Item("AR_CODE") = mAR_CODE.Text.trim
                    If AR_NAME.Text <> "" Then row.Item("AR_NAME_D") = AR_NAME.Text Else row.Item("AR_NAME_D") = "NEW"
                    row.Item("AR_NAME") = AR_NAME.Text
                    row.Item("AR_NAME_CH") = AR_NAME_CH.Text
                    row.Item("AR_X") = gU.decodeEmptyCdbl(AR_X.Text, 0)
                    row.Item("AR_Y") = gU.decodeEmptyCdbl(AR_Y.Text, 0)
                    row.Item("AR_LENGTH") = gU.decodeEmptyCdbl(AR_LENGTH.Text, 0)
                    row.Item("AR_WIDTH") = gU.decodeEmptyCdbl(AR_WIDTH.Text, 0)
                    row.Item("AR_HEIGHT") = gU.decodeEmptyCdbl(AR_HEIGHT.Text, 0)
                    row.Item("AR_GROSS_AREA") = gU.decodeEmptyCdbl(AR_GROSS_AREA.Text, 0)
                    row.Item("AR_GROSS_CBM") = gU.decodeEmptyCdbl(AR_GROSS_CBM.Text, 0)
                    row.Item("AR_FACILITIES") = AR_FACILITIES.Text
                    row.Item("AR_TEMP_FR") = AR_TEMP_FR.Text
                    row.Item("AR_TEMP_TO") = AR_TEMP_TO.Text

                    row.Item("AR_SECURITY") = AR_SECURITY.SelectedValue

                    'If AR_DAMAGE_YN.Checked Then
                    '    row.Item("AR_DAMAGE_YN") = "Y"
                    'Else
                    '    row.Item("AR_DAMAGE_YN") = "N"
                    'End If

                    If AR_POWER.Checked Then
                        row.Item("AR_POWER") = "Y"
                    Else
                        row.Item("AR_POWER") = "N"
                    End If

                    'If AR_TEMP_YN.Checked Then
                    '    row.Item("AR_TEMP_YN") = "Y"
                    'Else
                    '    row.Item("AR_TEMP_YN") = "N"
                    'End If


                    If AR_SCRAP_AREA.Checked Then row.Item("AR_SCRAP_AREA") = "Y" Else row.Item("AR_SCRAP_AREA") = "N"
                    'If AR_SHORTLEN_CABLE_AREA.Checked Then row.Item("AR_SHORTLEN_CABLE_AREA") = "Y" Else row.Item("AR_SHORTLEN_CABLE_AREA") = "N"
                    If AR_INSP_AREA.Checked Then row.Item("AR_INSP_AREA") = "Y" Else row.Item("AR_INSP_AREA") = "N"
                    If AR_REPAIR_AREA.Checked Then row.Item("AR_REPAIR_AREA") = "Y" Else row.Item("AR_REPAIR_AREA") = "N"
                    'If AR_PICKDROP_AREA.Checked Then row.Item("AR_PICKDROP_AREA") = "Y" Else row.Item("AR_PICKDROP_AREA") = "N"
                    If AR_FFI.Checked Then row.Item("AR_FFI") = "Y" Else row.Item("AR_FFI") = "N"

                    If AR_PICTURE_UPLOAD.HasFile Then
                        If Not arPhoto.ContainsKey(row.Item("AR_CODE").ToString) Then
                            arPhoto.Add(row.Item("AR_CODE").ToString, AR_PICTURE_UPLOAD.PostedFile)
                            row.Item("hasFile") = "Y"
                            'Session("arPhoto") = arPhoto
                        Else
                            Dim tempHttpPostFile As HttpPostedFile = arPhoto.Item(row.Item("AR_CODE").ToString)

                            If tempHttpPostFile.FileName <> AR_PICTURE_UPLOAD.PostedFile.FileName Then
                                arPhoto.Item(row.Item("AR_CODE").ToString) = AR_PICTURE_UPLOAD.PostedFile
                                'Session("arPhoto") = arPhoto
                            End If

                            tempHttpPostFile = Nothing
                        End If
                    End If
                End If
            Case 3
                If row IsNot Nothing Then
                    row.Item("remainYN") = ""
                    ViewState("notAllowAdd") = False

                    If mRK_CODE.Text.Trim <> "" Then
                        If RK_NAME.Text <> "" Then
                            If mRK_CODE.Text.Trim <> row.Item("RK_CODE").ToString Then
                                Dim dicCode As String = row.Item("FL_NUM").ToString & "||" & row.Item("AR_CODE").ToString & "||" & mRK_CODE.Text.Trim

                                If Not arDict.ContainsKey(dicCode) Then
                                    tblCode = row.Item("FL_NUM").ToString & "||" & row.Item("AR_CODE").ToString & "||" & row.Item("RK_CODE").ToString
                                    row.Item("RK_CODE") = mRK_CODE.Text.Trim
                                    row.Item("newPK") = ""
                                Else
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Duplicate Rack has found!!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                    End If

                                    row.Item("remainYN") = "Y"
                                    ViewState("notAllowAdd") = True

                                    Return False
                                End If
                            End If
                        Else
                            row.Item("remainYN") = "Y"
                            ViewState("notAllowAdd") = True

                            Return False
                        End If
                    Else
                        row.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True

                        Return False
                    End If

                    'row.Item("AR_CODE") = mAR_CODE.Text.trim
                    If RK_NAME.Text <> "" Then row.Item("RK_NAME_D") = RK_NAME.Text Else row.Item("RK_NAME_D") = "NEW"
                    row.Item("RK_NAME") = RK_NAME.Text
                    row.Item("RK_NAME_CH") = RK_NAME_CH.Text
                    row.Item("RK_GROSS_AREA") = gU.decodeEmptyCdbl(RK_GROSS_AREA.Text, 0)
                    row.Item("RK_GROSS_CBM") = gU.decodeEmptyCdbl(RK_GROSS_CBM.Text, 0)
                    row.Item("RK_X") = gU.decodeEmptyCdbl(RK_X.Text, 0)
                    row.Item("RK_Y") = gU.decodeEmptyCdbl(RK_Y.Text, 0)
                    row.Item("RK_LENGTH") = gU.decodeEmptyCdbl(RK_LENGTH.Text, 0)
                    row.Item("RK_WIDTH") = gU.decodeEmptyCdbl(RK_WIDTH.Text, 0)
                    row.Item("RK_DEPTH") = gU.decodeEmptyCdbl(RK_DEPTH.Text, 0)
                    row.Item("RK_XBINS") = gU.decodeEmptyCdbl(RK_XBINS.Text, 0)
                    row.Item("RK_YBINS") = gU.decodeEmptyCdbl(RK_YBINS.Text, 0)
                    row.Item("RK_REM") = RK_REM.Text
                End If
            Case 4
                If row IsNot Nothing Then
                    row.Item("remainYN") = ""
                    ViewState("notAllowAdd") = False

                    If mBN_CODE.Text.Trim <> "" Then
                        If mBN_CODE.Text.Trim <> row.Item("BN_CODE").ToString Then
                            Dim dicCode As String = row.Item("FL_NUM").ToString & "||" & row.Item("AR_CODE").ToString & "||" & row.Item("RK_CODE").ToString & "||" & mBN_CODE.Text.Trim

                            If Not bnDict.ContainsKey(dicCode) Then
                                tblCode = row.Item("FL_NUM").ToString & "||" & row.Item("AR_CODE").ToString & "||" & row.Item("RK_CODE").ToString & "||" & row.Item("BN_CODE").ToString
                                row.Item("BN_CODE") = mBN_CODE.Text.Trim
                                row.Item("newPK") = ""
                            Else
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Duplicate Bin has found!!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                End If

                                row.Item("remainYN") = "Y"
                                ViewState("notAllowAdd") = True

                                Return False
                            End If
                        End If
                    Else
                        row.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True

                        Return False
                    End If

                    If mBN_CODE.Text.Trim <> "" Then row.Item("BN_CODE_D") = mBN_CODE.Text.Trim 'Else row.Item("BN_CODE_D") = "NEW"
                    row.Item("BN_X") = gU.decodeEmptyCdbl(BN_X.Text, 0)
                    row.Item("BN_Y") = gU.decodeEmptyCdbl(BN_Y.Text, 0)
                    row.Item("BN_LENGTH") = gU.decodeEmptyCdbl(BN_LENGTH.Text, 0)
                    row.Item("BN_WIDTH") = gU.decodeEmptyCdbl(BN_WIDTH.Text, 0)
                    row.Item("BN_DEPTH") = gU.decodeEmptyCdbl(BN_DEPTH.Text, 0)
                    row.Item("BN_REM") = BN_REM.Text
                    row.Item("BN_UTILIZATION_TYPE") = BN_UTILIZATION_TYPE.SelectedValue
                    row.Item("BN_STATUS") = BN_STATUS.SelectedValue
                    row.Item("BN_CBM") = gU.decodeEmptyCdbl(BN_CBM.Text, 0)
                    row.Item("BN_CSMS_CODE") = BN_CSMS_CODE.Text

                End If
        End Select

        If tblCode <> "" Then old_Code = tblCode
        StoreData = True
    End Function

    Private Function StoreData(ByVal pFlcode As String, ByVal pArCode As String, ByVal pRkCode As String, ByVal pBnCode As String) As Boolean
        Dim tempCodeStr As String = ""
        Dim errorRaise As Boolean = False
        If pFlcode <> "" Then
            tempCodeStr = pFlcode
            Dim foundRow As DataRow = flDt.Rows.Find(tempCodeStr)

            If foundRow IsNot Nothing Then
                'foundRow.Item("FL_NUM") = gU.decodeEmptyCInt(mFL_NUM.Text.trim, 0)
                foundRow.Item("remainYN") = ""
                ViewState("notAllowAdd") = False

                If mFL_NUM.Text.Trim <> "" Then
                    If Not gU.isAlphaNumueric(mFL_NUM.Text.Trim) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Only Alphanumeric Floor Code is accepted!!", Session("gLang"))
                        End If

                        foundRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True
                        errorRaise = True

                        Return False
                    End If

                    If FL_NAME.Text <> "" Then
                        If foundRow.Item("mFlag").ToString = "N" Then
                            If mFL_NUM.Text.Trim <> foundRow.Item("FL_NUM").ToString Then
                                Dim nDataRows() As DataRow = flDt.Select("WH_CODE='" & WH_CODE.Value & "' AND FL_NUM = '" & mFL_NUM.Text.Trim & "'")

                                If nDataRows.Count = 0 Then
                                    foundRow.Item("FL_NUM") = If(mFL_NUM.Text.Trim <> "", mFL_NUM.Text.Trim, "00")
                                    foundRow.Item("newPK") = ""
                                    FL_NUM.Value = foundRow.Item("FL_NUM").ToString
                                Else
                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Duplicate Floor has found!!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                    End If

                                    foundRow.Item("remainYN") = "Y"
                                    ViewState("notAllowAdd") = True

                                    Return False
                                End If
                            End If
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Floor Name Cannot Be Empty!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Floor Name Cannot be Empty!!", Session("gLang"))
                        End If

                        foundRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True

                        errorRaise = True

                        Return False
                    End If
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Floor Cannot Be Empty!!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "Floor Cannot Be Empry!!", Session("gLang"))
                    End If

                    foundRow.Item("remainYN") = "Y"
                    ViewState("notAllowAdd") = True

                    errorRaise = True
                    Return False
                End If


                If FL_NAME.Text <> "" Then foundRow.Item("FL_NAME_D") = FL_NAME.Text Else foundRow.Item("FL_NAME_D") = "NEW"
                foundRow.Item("FL_NAME") = FL_NAME.Text
                foundRow.Item("FL_NAME_CH") = FL_NAME_CH.Text
                foundRow.Item("FL_GROSS_AREA") = gU.decodeEmptyCdbl(FL_GROSS_AREA.Text, 0)
                foundRow.Item("FL_GROSS_CBM") = gU.decodeEmptyCdbl(FL_GROSS_CBM.Text, 0)
                foundRow.Item("FL_MAX_WEIGHT") = gU.decodeEmptyCdbl(FL_MAX_WEIGHT.Text, 0)
                foundRow.Item("FL_LENGTH") = gU.decodeEmptyCdbl(FL_LENGTH.Text, 0)
                foundRow.Item("FL_WIDTH") = gU.decodeEmptyCdbl(FL_WIDTH.Text, 0)
                foundRow.Item("FL_HEIGHT") = gU.decodeEmptyCdbl(FL_HEIGHT.Text, 0)

                If FL_PICTURE_upload.HasFile Then
                    If Not flPhoto.ContainsKey(foundRow.Item("FL_NUM").ToString) Then
                        flPhoto.Add(foundRow.Item("FL_NUM").ToString, FL_PICTURE_upload.PostedFile)
                        foundRow.Item("hasFile") = "Y"
                        'Session("flPhoto") = flPhoto
                    Else
                        Dim tempHttpPostFile As HttpPostedFile = flPhoto.Item(foundRow.Item("FL_NUM").ToString)

                        If tempHttpPostFile.FileName <> FL_PICTURE_upload.PostedFile.FileName Then
                            flPhoto.Item(foundRow.Item("FL_NUM").ToString) = FL_PICTURE_upload.PostedFile
                            'Session("flPhoto") = flPhoto
                        End If

                        tempHttpPostFile = Nothing
                    End If
                End If
            End If

            flDt.AcceptChanges()
        End If

        If pArCode <> "" Then
            If tempCodeStr <> "" Then tempCodeStr = tempCodeStr & "||"
            tempCodeStr = tempCodeStr & pArCode

            Dim arRow As DataRow

            If arDict.ContainsKey(tempCodeStr) Then
                arRow = arDict.Item(tempCodeStr)

                arRow.Item("remainYN") = ""
                If errorRaise = False Then ViewState("notAllowAdd") = False

                If mAR_CODE.Text.Trim <> "" Then
                    If Not gU.isAlphaNumueric(mAR_CODE.Text.Trim) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Only Alphanumeric Area Code is accepted!!", Session("gLang"))
                        End If

                        arRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True
                        errorRaise = True

                        Return False
                    End If

                    If AR_NAME.Text <> "" Then
                        If arRow.Item("mFlag").ToString = "N" Then
                            If mAR_CODE.Text.Trim <> arRow.Item("AR_CODE").ToString Then
                                If Not arDict.ContainsKey(pFlcode & "||" & mAR_CODE.Text.Trim) Then
                                    arRow.Item("AR_CODE") = mAR_CODE.Text.Trim
                                    arRow.Item("newPK") = ""
                                    AR_CODE.Value = arRow.Item("AR_CODE").ToString
                                    arDict.Remove(pFlcode & "||" & pArCode)
                                    arDict.Add(pFlcode & "||" & mAR_CODE.Text.Trim, arRow)
                                End If
                            Else
                                'If Session("gLang") = "E" Then
                                '    uiFun.displayMsg(Me, "", "Duplicate Area has found!!", Session("gLang"))
                                'Else
                                '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                'End If

                                'arRow.Item("remainYN") = "Y"
                                'ViewState("notAllowAdd") = True
                            End If
                        Else
                            arRow.Item("AR_CODE") = mAR_CODE.Text.Trim
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Area Name Cannot Be Empty!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                        End If

                        arRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True
                        errorRaise = True

                        Return False
                    End If
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Area Code Cannot Be Empty!!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                    End If

                    arRow.Item("remainYN") = "Y"
                    ViewState("notAllowAdd") = True
                    errorRaise = True

                    Return False
                End If

                If AR_TEMP_FR.Text.ToString.Trim <> "" AndAlso Not IsNumeric(AR_TEMP_FR.Text.ToString.Trim) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Temperature From should be valid number!!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "Temperature From should be valid number!!", Session("gLang"))
                    End If

                    arRow.Item("remainYN") = "Y"
                    ViewState("notAllowAdd") = True
                    errorRaise = True

                    Return False
                End If

                If AR_TEMP_TO.Text.ToString.Trim <> "" AndAlso Not IsNumeric(AR_TEMP_TO.Text.ToString.Trim) Then

                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Temperature To should be valid number!!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "Temperature To should be valid number!!", Session("gLang"))
                    End If

                    arRow.Item("remainYN") = "Y"
                    ViewState("notAllowAdd") = True
                    errorRaise = True

                    Return False
                End If


                If AR_NAME.Text <> "" Then arRow.Item("AR_NAME_D") = AR_NAME.Text Else arRow.Item("AR_NAME_D") = "NEW"
                arRow.Item("AR_NAME") = AR_NAME.Text
                arRow.Item("AR_NAME_CH") = AR_NAME_CH.Text
                arRow.Item("AR_X") = gU.decodeEmptyCdbl(AR_X.Text, 0)
                arRow.Item("AR_Y") = gU.decodeEmptyCdbl(AR_Y.Text, 0)
                arRow.Item("AR_LENGTH") = gU.decodeEmptyCdbl(AR_LENGTH.Text, 0)
                arRow.Item("AR_WIDTH") = gU.decodeEmptyCdbl(AR_WIDTH.Text, 0)
                arRow.Item("AR_HEIGHT") = gU.decodeEmptyCdbl(AR_HEIGHT.Text, 0)
                arRow.Item("AR_GROSS_AREA") = gU.decodeEmptyCdbl(AR_GROSS_AREA.Text, 0)
                arRow.Item("AR_GROSS_CBM") = gU.decodeEmptyCdbl(AR_GROSS_CBM.Text, 0)
                arRow.Item("AR_FACILITIES") = AR_FACILITIES.Text
                arRow.Item("AR_TEMP_FR") = AR_TEMP_FR.Text
                arRow.Item("AR_TEMP_TO") = AR_TEMP_TO.Text

                arRow.Item("AR_SECURITY") = AR_SECURITY.SelectedValue

                'If AR_DAMAGE_YN.Checked Then
                '    arRow.Item("AR_DAMAGE_YN") = "Y"
                'Else
                '    arRow.Item("AR_DAMAGE_YN") = "N"
                'End If

                If AR_POWER.Checked Then
                    arRow.Item("AR_POWER") = "Y"
                Else
                    arRow.Item("AR_POWER") = "N"
                End If

                'If AR_TEMP_YN.Checked Then
                '    arRow.Item("AR_TEMP_YN") = "Y"
                'Else
                '    arRow.Item("AR_TEMP_YN") = "N"
                'End If

                If AR_SCRAP_AREA.Checked Then arRow.Item("AR_SCRAP_AREA") = "Y" Else arRow.Item("AR_SCRAP_AREA") = "N"
                'If AR_SHORTLEN_CABLE_AREA.Checked Then arRow.Item("AR_SHORTLEN_CABLE_AREA") = "Y" Else arRow.Item("AR_SHORTLEN_CABLE_AREA") = "N"
                If AR_INSP_AREA.Checked Then arRow.Item("AR_INSP_AREA") = "Y" Else arRow.Item("AR_INSP_AREA") = "N"
                If AR_REPAIR_AREA.Checked Then arRow.Item("AR_REPAIR_AREA") = "Y" Else arRow.Item("AR_REPAIR_AREA") = "N"
                'If AR_PICKDROP_AREA.Checked Then arRow.Item("AR_PICKDROP_AREA") = "Y" Else arRow.Item("AR_PICKDROP_AREA") = "N"
                If AR_FFI.Checked Then arRow.Item("AR_FFI") = "Y" Else arRow.Item("AR_FFI") = "N"


                If AR_PICTURE_UPLOAD.HasFile Then
                    If Not arPhoto.ContainsKey(arRow.Item("AR_CODE").ToString) Then
                        arPhoto.Add(arRow.Item("AR_CODE").ToString, AR_PICTURE_UPLOAD.PostedFile)
                        arRow.Item("hasFile") = "Y"
                        'Session("arPhoto") = arPhoto
                    Else
                        Dim tempHttpPostFile As HttpPostedFile = arPhoto.Item(arRow.Item("AR_CODE").ToString)

                        If tempHttpPostFile.FileName <> AR_PICTURE_UPLOAD.PostedFile.FileName Then
                            arPhoto.Item(arRow.Item("AR_CODE").ToString) = AR_PICTURE_UPLOAD.PostedFile
                            'Session("arPhoto") = arPhoto
                        End If

                        tempHttpPostFile = Nothing
                    End If
                End If

                'arRow.Item("AR_PICTURE") = AR_PICTURE.Text
            End If
        End If

        If pRkCode <> "" Then
            If tempCodeStr <> "" Then tempCodeStr = tempCodeStr & "||"
            tempCodeStr = tempCodeStr & pRkCode

            Dim rkRow As DataRow

            If rkDict.ContainsKey(tempCodeStr) Then
                rkRow = rkDict.Item(tempCodeStr)

                rkRow.Item("remainYN") = ""
                If errorRaise = False Then ViewState("notAllowAdd") = False

                If mRK_CODE.Text.Trim <> "" Then
                    If Not gU.isAlphaNumueric(mRK_CODE.Text.Trim) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Only Alphanumeric Rack Code is accepted!!", Session("gLang"))
                        End If

                        rkRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True
                        errorRaise = True

                        Return False
                    End If

                    If RK_NAME.Text <> "" Then
                        If rkRow.Item("mFlag").ToString = "N" Then
                            If mRK_CODE.Text.Trim <> rkRow.Item("RK_CODE").ToString Then
                                If Not rkDict.ContainsKey(pFlcode & "||" & pArCode & "||" & mRK_CODE.Text.Trim) Then
                                    rkRow.Item("RK_CODE") = mRK_CODE.Text.Trim
                                    rkRow.Item("newPK") = ""
                                    'AR_CODE.Value = rkRow.Item("AR_CODE").ToString
                                    RK_CODE.Value = rkRow.Item("RK_CODE").ToString
                                    rkDict.Remove(pFlcode & "||" & pArCode & "||" & pRkCode)
                                    rkDict.Add(pFlcode & "||" & pArCode & "||" & mRK_CODE.Text.Trim, rkRow)
                                End If
                            Else
                                'If Session("gLang") = "E" Then
                                '    uiFun.displayMsg(Me, "", "Duplicate Rack has found!!", Session("gLang"))
                                'Else
                                '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                                'End If

                                'rkRow.Item("remainYN") = "Y"
                                'ViewState("notAllowAdd") = True
                            End If
                        Else
                            rkRow.Item("RK_CODE") = mRK_CODE.Text.Trim
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Rack Name Cannot Be Empty!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                        End If

                        rkRow.Item("remainYN") = "Y"
                        ViewState("notAllowAdd") = True
                        errorRaise = True

                        Return False
                    End If
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Rack Code Cannot Be Empty!!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                    End If

                    rkRow.Item("remainYN") = "Y"
                    ViewState("notAllowAdd") = True
                    errorRaise = True

                    Return False
                End If

                If RK_NAME.Text <> "" Then rkRow.Item("RK_NAME_D") = RK_NAME.Text Else rkRow.Item("RK_NAME_D") = "NEW"
                rkRow.Item("RK_NAME") = RK_NAME.Text
                rkRow.Item("RK_NAME_CH") = RK_NAME_CH.Text
                rkRow.Item("RK_GROSS_AREA") = gU.decodeEmptyCdbl(RK_GROSS_AREA.Text, 0)
                rkRow.Item("RK_GROSS_CBM") = gU.decodeEmptyCdbl(RK_GROSS_CBM.Text, 0)
                rkRow.Item("RK_X") = gU.decodeEmptyCdbl(RK_X.Text, 0)
                rkRow.Item("RK_Y") = gU.decodeEmptyCdbl(RK_Y.Text, 0)
                rkRow.Item("RK_LENGTH") = gU.decodeEmptyCdbl(RK_LENGTH.Text, 0)
                rkRow.Item("RK_WIDTH") = gU.decodeEmptyCdbl(RK_WIDTH.Text, 0)
                rkRow.Item("RK_DEPTH") = gU.decodeEmptyCdbl(RK_DEPTH.Text, 0)
                rkRow.Item("RK_XBINS") = gU.decodeEmptyCdbl(RK_XBINS.Text, 0)
                rkRow.Item("RK_YBINS") = gU.decodeEmptyCdbl(RK_YBINS.Text, 0)
                rkRow.Item("RK_REM") = RK_REM.Text
            End If
        End If

        If pBnCode <> "" Then
            If tempCodeStr <> "" Then tempCodeStr = tempCodeStr & "||"
            tempCodeStr = tempCodeStr & pBnCode

            Dim bnRow As DataRow

            If bnDict.ContainsKey(tempCodeStr) Then
                bnRow = bnDict.Item(tempCodeStr)

                bnRow.Item("remainYN") = ""
                If errorRaise = False Then ViewState("notAllowAdd") = False

                If mBN_CODE.Text.Trim <> "" Then
                    If bnRow.Item("mFlag").ToString = "N" Then
                        If mBN_CODE.Text.Trim <> bnRow.Item("BN_CODE").ToString Then
                            If Not bnDict.ContainsKey(pFlcode & "||" & pArCode & "||" & pRkCode & "||" & mBN_CODE.Text.Trim) Then
                                bnRow.Item("BN_CODE") = mBN_CODE.Text.Trim
                                bnRow.Item("newPK") = ""
                                'AR_CODE.Value = rkRow.Item("AR_CODE").ToString
                                BN_CODE.Value = bnRow.Item("BN_CODE").ToString
                                bnDict.Remove(pFlcode & "||" & pArCode & "||" & pRkCode & "||" & pBnCode)
                                bnDict.Add(pFlcode & "||" & pArCode & "||" & pRkCode & "||" & mBN_CODE.Text.Trim, bnRow)
                            End If
                        Else
                            'If Session("gLang") = "E" Then
                            '    uiFun.displayMsg(Me, "", "Duplicate Rack has found!!", Session("gLang"))
                            'Else
                            '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                            'End If

                            'rkRow.Item("remainYN") = "Y"
                            'ViewState("notAllowAdd") = True
                        End If
                    Else
                        bnRow.Item("BN_CODE") = mBN_CODE.Text.Trim
                    End If
                Else
                    'If Session("gLang") = "E" Then
                    '    uiFun.displayMsg(Me, "", "Bin Code Cannot Be Empty!!", Session("gLang"))
                    'Else
                    '    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                    'End If

                    'bnRow.Item("remainYN") = "Y"
                    'ViewState("notAllowAdd") = True
                    'errorRaise = True
                End If

                If mBN_CODE.Text.Trim <> "" Then bnRow.Item("BN_CODE_D") = mBN_CODE.Text.Trim 'Else bnRow.Item("BN_CODE_D") = "NEW"
                bnRow.Item("BN_X") = gU.decodeEmptyCdbl(BN_X.Text, 0)
                bnRow.Item("BN_Y") = gU.decodeEmptyCdbl(BN_Y.Text, 0)
                bnRow.Item("BN_LENGTH") = gU.decodeEmptyCdbl(BN_LENGTH.Text, 0)
                bnRow.Item("BN_WIDTH") = gU.decodeEmptyCdbl(BN_WIDTH.Text, 0)
                bnRow.Item("BN_DEPTH") = gU.decodeEmptyCdbl(BN_DEPTH.Text, 0)
                bnRow.Item("BN_REM") = BN_REM.Text
                bnRow.Item("BN_UTILIZATION_TYPE") = BN_UTILIZATION_TYPE.SelectedValue
                bnRow.Item("BN_STATUS") = BN_STATUS.SelectedValue
                bnRow.Item("BN_CBM") = gU.decodeEmptyCdbl(BN_CBM.Text, 0)
                bnRow.Item("BN_CSMS_CODE") = BN_CSMS_CODE.Text
            End If
        End If

        Return True

    End Function

    Private Sub ShowTable(ByVal type As Integer)
        MainTbl.Visible = True

        Select Case type
            Case 1
                flTr.Visible = True
                arTr.Visible = False
                rkTr.Visible = False
                bnTr.Visible = False

                If Session("gLang") = "E" Then
                    lheader.Text = "Floor Maintenance"
                ElseIf Session("gLang") = "C" Then
                    lheader.Text = ""
                End If

                Dim flRows() As DataRow
                flRows = flDt.Select("FL_NUM = '" & FL_NUM.Value & "'")

                If flRows.Count > 0 Then
                    If flRows(0).Item("mFlag").ToString = "N" Then

                        mFL_NUM.CssClass = "REQUIRED"
                        FL_NAME.CssClass = "REQUIRED"

                        If flRows(0).Item("remainYN").ToString <> "Y" Then
                            If flRows(0).Item("newPK").ToString = "Y" Then
                                mFL_NUM.Text = ""
                            Else
                                mFL_NUM.Text = flRows(0).Item("FL_NUM").ToString
                            End If
                        End If

                        Dim pic_preview_yn As Boolean = False

                        FL_PICTURE.Visible = True
                        FL_PICTURE.Enabled = True
                        FL_PICTURE_upload.Visible = True
                        FL_PICTURE_upload.Enabled = True
                        FL_PICTURE_edit.Visible = False

                        FL_PICTURE.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & flRows(0).Item("FL_PICTURE").ToString & "&refresh=true&Height=180&code=MAST_WMC&userid=" & Session("usr_id")

                        If flRows(0).Item("removeFile").ToString = "Y" Then
                            FL_PICTURE_remove.Visible = True
                            FL_PICTURE_remove.Enabled = False
                            FL_PICTURE.Enabled = False
                            fl_preview_pic.Enabled = False
                            FL_PICTURE_edit.Visible = False
                            FL_PICTURE_lit.Text = ""
                            FL_PICTURE_lit.Visible = False
                            FL_PICTURE_upload.Enabled = False

                            pic_preview_yn = True

                        ElseIf flRows(0).Item("hasFile").ToString = "Y" Then
                            If flPhoto.ContainsKey(flRows(0).Item("FL_NUM").ToString) Then
                                Dim tempPostedFile As HttpPostedFile = flPhoto.Item(flRows(0).Item("FL_NUM").ToString)

                                FL_PICTURE_lit.Text = tempPostedFile.FileName
                                FL_PICTURE_upload.Visible = False
                                FL_PICTURE_lit.Visible = True
                                FL_PICTURE_edit.Visible = True

                                pic_preview_yn = False
                            End If
                        Else
                            If flRows(0).Item("FL_PICTURE").ToString <> "" Then
                                FL_PICTURE.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & flRows(0).Item("FL_PICTURE").ToString & "&refresh=true&ds=true&code=MAST_WM&userid=" & Session("usr_id")
                                FL_PICTURE.Target = "_new1"

                                fl_preview_pic.Visible = True
                                FL_PICTURE_remove.Visible = True
                                FL_PICTURE.Enabled = True
                                pic_preview_yn = True
                            Else
                                fl_preview_pic.Visible = False
                                FL_PICTURE_remove.Visible = False
                                FL_PICTURE.Enabled = False
                            End If
                        End If

                        If pic_preview_yn Then
                            fl_pic_tr.Visible = True
                            fl_remove_tr.Visible = True
                        Else
                            fl_pic_tr.Visible = False
                            fl_remove_tr.Visible = False
                        End If

                        FL_NAME.Text = flRows(0).Item("FL_NAME").ToString
                        FL_NAME_CH.Text = flRows(0).Item("FL_NAME_CH").ToString
                        FL_GROSS_AREA.Text = cU.FormatDecimalwString(flRows(0).Item("FL_GROSS_AREA").ToString)
                        FL_GROSS_CBM.Text = cU.FormatDecimalwString(flRows(0).Item("FL_GROSS_CBM").ToString)
                        FL_MAX_WEIGHT.Text = cU.FormatDecimalwString(flRows(0).Item("FL_MAX_WEIGHT").ToString)
                        FL_LENGTH.Text = cU.FormatDecimalwString(flRows(0).Item("FL_LENGTH").ToString)
                        FL_WIDTH.Text = cU.FormatDecimalwString(flRows(0).Item("FL_WIDTH").ToString)
                        FL_HEIGHT.Text = cU.FormatDecimalwString(flRows(0).Item("FL_HEIGHT").ToString)

                        sys_cb.Text = flRows(0).Item("sys_cb").ToString
                        sys_lub.Text = flRows(0).Item("sys_lub").ToString
                        sys_cd.Text = cU.chgToFullDF(flRows(0).Item("sys_cd").ToString)
                        sys_lud.Text = cU.chgToFullDF(flRows(0).Item("sys_lud").ToString)

                        mFL_NUM.BorderWidth = Nothing
                        mFL_NUM.BackColor = Drawing.Color.Empty
                        mFL_NUM.ReadOnly = False
                    Else
                        Dim pic_preview_yn As Boolean = False

                        FL_PICTURE.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & flRows(0).Item("FL_PICTURE").ToString & "&refresh=true&Height=180&code=MAST_WM&userid=" & Session("usr_id")

                        FL_PICTURE.Visible = True
                        FL_PICTURE.Enabled = True
                        FL_PICTURE_upload.Visible = True
                        FL_PICTURE_upload.Enabled = True
                        FL_PICTURE_edit.Visible = False

                        If flRows(0).Item("removeFile").ToString = "Y" Then
                            FL_PICTURE_remove.Visible = True
                            FL_PICTURE_remove.Enabled = False
                            FL_PICTURE.Enabled = False
                            fl_preview_pic.Enabled = False
                            FL_PICTURE_edit.Visible = False
                            FL_PICTURE_lit.Text = ""
                            FL_PICTURE_lit.Visible = False
                            FL_PICTURE_upload.Enabled = False

                            pic_preview_yn = True

                        ElseIf flRows(0).Item("hasFile").ToString = "Y" Then
                            If flPhoto.ContainsKey(flRows(0).Item("FL_NUM").ToString) Then
                                Dim tempPostedFile As HttpPostedFile = flPhoto.Item(flRows(0).Item("FL_NUM").ToString)

                                FL_PICTURE_lit.Text = tempPostedFile.FileName
                                FL_PICTURE_upload.Visible = False
                                FL_PICTURE_lit.Visible = True
                                FL_PICTURE_edit.Visible = True
                                'FL_PICTURE.Visible = true

                                pic_preview_yn = False
                            End If
                        Else
                            If flRows(0).Item("FL_PICTURE").ToString <> "" Then
                                FL_PICTURE.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & flRows(0).Item("FL_PICTURE").ToString & "&refresh=true&ds=true&code=MAST_WM&userid=" & Session("usr_id")
                                FL_PICTURE.Target = "_new1"

                                FL_PICTURE_lit.Text = ""
                                FL_PICTURE_lit.Visible = False
                                fl_preview_pic.Visible = True
                                fl_preview_pic.Enabled = True
                                FL_PICTURE_remove.Visible = True
                                FL_PICTURE_remove.Enabled = True
                                pic_preview_yn = True
                            Else

                                FL_PICTURE_lit.Text = ""
                                FL_PICTURE_lit.Visible = False
                                fl_preview_pic.Visible = False
                                FL_PICTURE_remove.Visible = False
                                FL_PICTURE.Enabled = False
                            End If
                        End If

                        If pic_preview_yn Then
                            fl_pic_tr.Visible = True
                            fl_remove_tr.Visible = True
                        Else
                            fl_pic_tr.Visible = False
                            fl_remove_tr.Visible = False
                        End If

                        mFL_NUM.Text = flRows(0).Item("FL_NUM").ToString
                        FL_NAME.Text = flRows(0).Item("FL_NAME").ToString
                        FL_NAME_CH.Text = flRows(0).Item("FL_NAME_CH").ToString
                        FL_GROSS_AREA.Text = cU.FormatDecimalwString(flRows(0).Item("FL_GROSS_AREA").ToString)
                        FL_GROSS_CBM.Text = cU.FormatDecimalwString(flRows(0).Item("FL_GROSS_CBM").ToString)
                        FL_MAX_WEIGHT.Text = cU.FormatDecimalwString(flRows(0).Item("FL_MAX_WEIGHT").ToString)
                        FL_LENGTH.Text = cU.FormatDecimalwString(flRows(0).Item("FL_LENGTH").ToString)
                        FL_WIDTH.Text = cU.FormatDecimalwString(flRows(0).Item("FL_WIDTH").ToString)
                        FL_HEIGHT.Text = cU.FormatDecimalwString(flRows(0).Item("FL_HEIGHT").ToString)

                        sys_cb.Text = flRows(0).Item("sys_cb").ToString
                        sys_lub.Text = flRows(0).Item("sys_lub").ToString
                        sys_cd.Text = cU.chgToFullDF(flRows(0).Item("sys_cd").ToString)
                        sys_lud.Text = cU.chgToFullDF(flRows(0).Item("sys_lud").ToString)

                        mFL_NUM.BorderWidth = 0
                        mFL_NUM.BackColor = Drawing.Color.Transparent
                        mFL_NUM.ReadOnly = True
                    End If
                End If
            Case 2
                flTr.Visible = False
                arTr.Visible = True
                rkTr.Visible = False
                bnTr.Visible = False

                uiFun.load_dropdown(AR_SECURITY, "Select COLC_CODE, COLC_ENG_VALUE from WMS_COL_CODE where COLC_TABCOL='WMS_WH_AREA.AR_SECURITY' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))

                If Session("gLang") = "E" Then
                    lheader.Text = "Area Maintenance"
                ElseIf Session("gLang") = "C" Then
                    lheader.Text = ""
                End If

                Dim arRow As DataRow

                arRow = arDict.Item(FL_NUM.Value & "||" & AR_CODE.Value)

                If arRow.Item("mFlag").ToString = "N" Then

                    mAR_CODE.CssClass = "REQUIRED"
                    AR_NAME.CssClass = "REQUIRED"

                    If arRow.Item("remainYN").ToString <> "Y" Then
                        If arRow.Item("newPK").ToString = "Y" Then
                            mAR_CODE.Text = ""
                        Else
                            mAR_CODE.Text = arRow.Item("AR_CODE").ToString
                        End If
                    End If

                    Dim pic_preview_yn As Boolean = False

                    AR_PICTURE.Visible = True
                    AR_PICTURE.Enabled = True
                    AR_PICTURE_UPLOAD.Visible = True
                    AR_PICTURE_UPLOAD.Enabled = True
                    AR_PICTURE_EDIT.Visible = False

                    AR_PICTURE.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & arRow.Item("AR_PICTURE").ToString & "&refresh=true&Height=180&code=MAST_WMC&userid=" & Session("usr_id")

                    If arRow.Item("removeFile").ToString = "Y" Then
                        AR_PICTURE_REMOVE.Visible = True
                        AR_PICTURE_REMOVE.Enabled = False
                        AR_PICTURE.Enabled = False
                        ar_preview_pic.Enabled = False
                        AR_PICTURE_EDIT.Visible = False
                        AR_PICTURE_LIT.Text = ""
                        AR_PICTURE_LIT.Visible = False
                        AR_PICTURE_UPLOAD.Enabled = False

                        pic_preview_yn = True

                    ElseIf arRow.Item("hasFile").ToString = "Y" Then
                        If flPhoto.ContainsKey(arRow.Item("ar_NUM").ToString) Then
                            Dim tempPostedFile As HttpPostedFile = flPhoto.Item(arRow.Item("ar_NUM").ToString)

                            AR_PICTURE_LIT.Text = tempPostedFile.FileName
                            AR_PICTURE_UPLOAD.Visible = False
                            AR_PICTURE_LIT.Visible = True
                            AR_PICTURE_EDIT.Visible = True

                            pic_preview_yn = False
                        End If
                    Else
                        If arRow.Item("ar_PICTURE").ToString <> "" Then
                            AR_PICTURE.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & arRow.Item("AR_PICTURE").ToString & "&refresh=true&ds=true&code=MAST_WM&userid=" & Session("usr_id")
                            AR_PICTURE.Target = "_new1"

                            ar_preview_pic.Visible = True
                            AR_PICTURE_REMOVE.Visible = True
                            AR_PICTURE.Enabled = True
                            pic_preview_yn = True
                        Else
                            ar_preview_pic.Visible = False
                            AR_PICTURE_REMOVE.Visible = False
                            AR_PICTURE.Enabled = False
                        End If
                    End If

                    If pic_preview_yn Then
                        ar_pic_tr.Visible = True
                        ar_remove_tr.Visible = True
                    Else
                        ar_pic_tr.Visible = False
                        ar_remove_tr.Visible = False
                    End If

                    AR_NAME.Text = arRow.Item("AR_NAME").ToString
                    AR_NAME_CH.Text = arRow.Item("AR_NAME_CH").ToString
                    AR_X.Text = cU.FormatDecimalwString(arRow.Item("AR_X").ToString)
                    AR_Y.Text = cU.FormatDecimalwString(arRow.Item("AR_Y").ToString)
                    AR_LENGTH.Text = cU.FormatDecimalwString(arRow.Item("AR_LENGTH").ToString)
                    AR_WIDTH.Text = cU.FormatDecimalwString(arRow.Item("AR_WIDTH").ToString)
                    AR_HEIGHT.Text = cU.FormatDecimalwString(arRow.Item("AR_HEIGHT").ToString)
                    AR_GROSS_AREA.Text = cU.FormatDecimalwString(arRow.Item("AR_GROSS_AREA").ToString)
                    AR_GROSS_CBM.Text = cU.FormatDecimalwString(arRow.Item("AR_GROSS_CBM").ToString)
                    AR_FACILITIES.Text = arRow.Item("AR_FACILITIES").ToString
                    'AR_PICTURE.text = 

                    AR_TEMP_FR.Text = arRow.Item("AR_TEMP_FR").ToString
                    AR_TEMP_TO.Text = arRow.Item("AR_TEMP_TO").ToString

                    AR_SECURITY.SelectedValue = arRow.Item("AR_SECURITY").ToString

                    'If arRow.Item("AR_DAMAGE_YN").ToString = "Y" Then
                    '    AR_DAMAGE_YN.Checked = True
                    'Else
                    '    AR_DAMAGE_YN.Checked = False
                    'End If

                    If arRow.Item("AR_POWER").ToString = "Y" Then
                        AR_POWER.Checked = True
                    Else
                        AR_POWER.Checked = False
                    End If

                    'If arRow.Item("AR_TEMP_YN").ToString = "Y" Then
                    '    AR_TEMP_YN.Checked = True
                    'Else
                    '    AR_TEMP_YN.Checked = False
                    'End If

                    If arRow.Item("AR_SCRAP_AREA").ToString.Trim = "Y" Then AR_SCRAP_AREA.Checked = True Else AR_SCRAP_AREA.Checked = False
                    'If arRow.Item("AR_SHORTLEN_CABLE_AREA").ToString.Trim = "Y" Then AR_SHORTLEN_CABLE_AREA.Checked = True Else AR_SHORTLEN_CABLE_AREA.Checked = False
                    If arRow.Item("AR_INSP_AREA").ToString.Trim = "Y" Then AR_INSP_AREA.Checked = True Else AR_INSP_AREA.Checked = False
                    If arRow.Item("AR_REPAIR_AREA").ToString.Trim = "Y" Then AR_REPAIR_AREA.Checked = True Else AR_REPAIR_AREA.Checked = False
                    'If arRow.Item("AR_PICKDROP_AREA").ToString.Trim = "Y" Then AR_PICKDROP_AREA.Checked = True Else AR_PICKDROP_AREA.Checked = False
                    AR_FFI.Checked = True

                    sys_cb.Text = arRow.Item("sys_cb").ToString
                    sys_lub.Text = arRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(arRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(arRow.Item("sys_lud").ToString)

                    mAR_CODE.BorderWidth = Nothing
                    mAR_CODE.BackColor = Drawing.Color.Empty
                    mAR_CODE.ReadOnly = False
                Else
                    Dim pic_preview_yn As Boolean = False

                    AR_PICTURE.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & arRow.Item("AR_PICTURE").ToString & "&refresh=true&Height=180&code=MAST_WM&userid=" & Session("usr_id")

                    AR_PICTURE.Visible = True
                    AR_PICTURE.Enabled = True
                    AR_PICTURE_UPLOAD.Visible = True
                    AR_PICTURE_UPLOAD.Enabled = True
                    AR_PICTURE_EDIT.Visible = False

                    If arRow.Item("removeFile").ToString = "Y" Then
                        AR_PICTURE_REMOVE.Visible = True
                        AR_PICTURE_REMOVE.Enabled = False
                        AR_PICTURE.Enabled = False
                        ar_preview_pic.Enabled = False
                        AR_PICTURE_EDIT.Visible = False
                        AR_PICTURE_LIT.Text = ""
                        AR_PICTURE_LIT.Visible = False
                        AR_PICTURE_UPLOAD.Enabled = False

                        pic_preview_yn = True

                    ElseIf arRow.Item("hasFile").ToString = "Y" Then
                        If flPhoto.ContainsKey(arRow.Item("ar_NUM").ToString) Then
                            Dim tempPostedFile As HttpPostedFile = flPhoto.Item(arRow.Item("ar_NUM").ToString)

                            AR_PICTURE_LIT.Text = tempPostedFile.FileName
                            AR_PICTURE_UPLOAD.Visible = False
                            AR_PICTURE_LIT.Visible = True
                            AR_PICTURE_EDIT.Visible = True
                            'ar_PICTURE.Visible = true

                            pic_preview_yn = False
                        End If
                    Else
                        If arRow.Item("ar_PICTURE").ToString <> "" Then
                            AR_PICTURE.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & arRow.Item("ar_PICTURE").ToString & "&refresh=true&ds=true&code=MAST_WM&userid=" & Session("usr_id")
                            AR_PICTURE.Target = "_new1"

                            AR_PICTURE_LIT.Text = ""
                            AR_PICTURE_LIT.Visible = False
                            ar_preview_pic.Visible = True
                            ar_preview_pic.Enabled = True
                            AR_PICTURE_REMOVE.Visible = True
                            AR_PICTURE_REMOVE.Enabled = True
                            pic_preview_yn = True
                        Else

                            AR_PICTURE_LIT.Text = ""
                            AR_PICTURE_LIT.Visible = False
                            ar_preview_pic.Visible = False
                            AR_PICTURE_REMOVE.Visible = False
                            AR_PICTURE.Enabled = False
                        End If
                    End If

                    If pic_preview_yn Then
                        ar_pic_tr.Visible = True
                        ar_remove_tr.Visible = True
                    Else
                        ar_pic_tr.Visible = False
                        ar_remove_tr.Visible = False
                    End If

                    mAR_CODE.Text = arRow.Item("AR_CODE").ToString
                    AR_NAME.Text = arRow.Item("AR_NAME").ToString
                    AR_NAME_CH.Text = arRow.Item("AR_NAME_CH").ToString
                    AR_X.Text = cU.FormatDecimalwString(arRow.Item("AR_X").ToString)
                    AR_Y.Text = cU.FormatDecimalwString(arRow.Item("AR_Y").ToString)
                    AR_LENGTH.Text = cU.FormatDecimalwString(arRow.Item("AR_LENGTH").ToString)
                    AR_WIDTH.Text = cU.FormatDecimalwString(arRow.Item("AR_WIDTH").ToString)
                    AR_HEIGHT.Text = cU.FormatDecimalwString(arRow.Item("AR_HEIGHT").ToString)
                    AR_GROSS_AREA.Text = cU.FormatDecimalwString(arRow.Item("AR_GROSS_AREA").ToString)
                    AR_GROSS_CBM.Text = cU.FormatDecimalwString(arRow.Item("AR_GROSS_CBM").ToString)
                    AR_FACILITIES.Text = arRow.Item("AR_FACILITIES").ToString
                    'AR_PICTURE.text = 

                    AR_TEMP_FR.Text = arRow.Item("AR_TEMP_FR").ToString
                    AR_TEMP_TO.Text = arRow.Item("AR_TEMP_TO").ToString

                    AR_SECURITY.SelectedValue = arRow.Item("AR_SECURITY").ToString

                    'If arRow.Item("AR_DAMAGE_YN").ToString = "Y" Then
                    '    AR_DAMAGE_YN.Checked = True
                    'Else
                    '    AR_DAMAGE_YN.Checked = False
                    'End If

                    If arRow.Item("AR_POWER").ToString = "Y" Then
                        AR_POWER.Checked = True
                    Else
                        AR_POWER.Checked = False
                    End If

                    'If arRow.Item("AR_TEMP_YN").ToString = "Y" Then
                    '    AR_TEMP_YN.Checked = True
                    'Else
                    '    AR_TEMP_YN.Checked = False
                    'End If

                    If arRow.Item("AR_SCRAP_AREA").ToString.Trim = "Y" Then AR_SCRAP_AREA.Checked = True Else AR_SCRAP_AREA.Checked = False
                    'If arRow.Item("AR_SHORTLEN_CABLE_AREA").ToString.Trim = "Y" Then AR_SHORTLEN_CABLE_AREA.Checked = True Else AR_SHORTLEN_CABLE_AREA.Checked = False
                    If arRow.Item("AR_INSP_AREA").ToString.Trim = "Y" Then AR_INSP_AREA.Checked = True Else AR_INSP_AREA.Checked = False
                    If arRow.Item("AR_REPAIR_AREA").ToString.Trim = "Y" Then AR_REPAIR_AREA.Checked = True Else AR_REPAIR_AREA.Checked = False
                    'If arRow.Item("AR_PICKDROP_AREA").ToString.Trim = "Y" Then AR_PICKDROP_AREA.Checked = True Else AR_PICKDROP_AREA.Checked = False
                    If arRow.Item("AR_FFI").ToString.Trim = "Y" Then AR_FFI.Checked = True Else AR_FFI.Checked = False


                    sys_cb.Text = arRow.Item("sys_cb").ToString
                    sys_lub.Text = arRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(arRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(arRow.Item("sys_lud").ToString)

                    mAR_CODE.BorderWidth = 0
                    mAR_CODE.BackColor = Drawing.Color.Transparent
                    mAR_CODE.ReadOnly = True
                End If

            Case 3
                flTr.Visible = False
                arTr.Visible = False
                rkTr.Visible = True
                bnTr.Visible = False

                If Session("gLang") = "E" Then
                    lheader.Text = "Rack Maintenance"
                ElseIf Session("gLang") = "C" Then
                    lheader.Text = ""
                End If

                Dim rkRow As DataRow

                rkRow = rkDict.Item(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value)

                If rkRow.Item("mFlag").ToString = "N" Then

                    mRK_CODE.CssClass = "REQUIRED"
                    RK_NAME.CssClass = "REQUIRED"

                    If rkRow.Item("remainYN").ToString <> "Y" Then
                        If rkRow.Item("newPK").ToString = "Y" Then
                            mRK_CODE.Text = ""
                        Else
                            mRK_CODE.Text = rkRow.Item("RK_CODE").ToString
                        End If
                    End If

                    RK_NAME.Text = rkRow.Item("RK_NAME").ToString
                    RK_NAME_CH.Text = rkRow.Item("RK_NAME_CH").ToString
                    RK_GROSS_AREA.Text = cU.FormatDecimalwString(rkRow.Item("RK_GROSS_AREA").ToString)
                    RK_GROSS_CBM.Text = cU.FormatDecimalwString(rkRow.Item("RK_GROSS_CBM").ToString)
                    RK_X.Text = cU.FormatIntegerString(rkRow.Item("RK_X").ToString)
                    RK_Y.Text = cU.FormatIntegerString(rkRow.Item("RK_Y").ToString)
                    RK_LENGTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_LENGTH").ToString)
                    RK_WIDTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_WIDTH").ToString)
                    RK_DEPTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_DEPTH").ToString)
                    RK_XBINS.Text = cU.FormatIntegerString(rkRow.Item("RK_XBINS").ToString)
                    RK_YBINS.Text = cU.FormatIntegerString(rkRow.Item("RK_YBINS").ToString)
                    RK_REM.Text = rkRow.Item("RK_REM").ToString

                    sys_cb.Text = rkRow.Item("sys_cb").ToString
                    sys_lub.Text = rkRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(rkRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(rkRow.Item("sys_lud").ToString)

                    mRK_CODE.BorderWidth = Nothing
                    mRK_CODE.BackColor = Drawing.Color.Empty
                    mRK_CODE.ReadOnly = False
                Else
                    mRK_CODE.Text = rkRow.Item("RK_CODE").ToString
                    RK_NAME.Text = rkRow.Item("RK_NAME").ToString
                    RK_NAME_CH.Text = rkRow.Item("RK_NAME_CH").ToString
                    RK_GROSS_AREA.Text = cU.FormatDecimalwString(rkRow.Item("RK_GROSS_AREA").ToString)
                    RK_GROSS_CBM.Text = cU.FormatDecimalwString(rkRow.Item("RK_GROSS_CBM").ToString)
                    RK_X.Text = cU.FormatIntegerString(rkRow.Item("RK_X").ToString)
                    RK_Y.Text = cU.FormatIntegerString(rkRow.Item("RK_Y").ToString)
                    RK_LENGTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_LENGTH").ToString)
                    RK_WIDTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_WIDTH").ToString)
                    RK_DEPTH.Text = cU.FormatDecimalwString(rkRow.Item("RK_DEPTH").ToString)
                    RK_XBINS.Text = cU.FormatIntegerString(rkRow.Item("RK_XBINS").ToString)
                    RK_YBINS.Text = cU.FormatIntegerString(rkRow.Item("RK_YBINS").ToString)
                    RK_REM.Text = rkRow.Item("RK_REM").ToString

                    sys_cb.Text = rkRow.Item("sys_cb").ToString
                    sys_lub.Text = rkRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(rkRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(rkRow.Item("sys_lud").ToString)

                    mRK_CODE.BorderWidth = 0
                    mRK_CODE.BackColor = Drawing.Color.Transparent
                    mRK_CODE.ReadOnly = True
                End If

            Case 4
                flTr.Visible = False
                arTr.Visible = False
                rkTr.Visible = False
                bnTr.Visible = True

                uiFun.load_dropdown(BN_STATUS, "select COLC_CODE, COLC_ENG_VALUE from WMS_COL_CODE where COLC_TABCOL='WMS_WH_BIN.BN_STATUS' order by colc_display_seq", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))

                If Session("gLang") = "E" Then
                    lheader.Text = "Bin Maintenance"
                ElseIf Session("gLang") = "C" Then
                    lheader.Text = ""
                End If

                Dim bnRow As DataRow

                bnRow = bnDict.Item(FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & BN_CODE.Value)

                BN_LENGTH.Attributes.Add("onkeyup", _
                        "document.forms[0]." & BN_CBM.ClientID & ".value = " & _
                        "fixDecimal((this.value * " & _
                        "document.forms[0]." & BN_WIDTH.ClientID & ".value * " & _
                        "document.forms[0]." & BN_DEPTH.ClientID & ".value) / 1000000,3);")

                BN_WIDTH.Attributes.Add("onkeyup", _
                        "document.forms[0]." & BN_CBM.ClientID & ".value = " & _
                        "fixDecimal((this.value * " & _
                        "document.forms[0]." & BN_LENGTH.ClientID & ".value * " & _
                        "document.forms[0]." & BN_DEPTH.ClientID & ".value) / 1000000,3);")

                BN_DEPTH.Attributes.Add("onkeyup", _
                        "document.forms[0]." & BN_CBM.ClientID & ".value = " & _
                        "fixDecimal((this.value * " & _
                        "document.forms[0]." & BN_WIDTH.ClientID & ".value * " & _
                        "document.forms[0]." & BN_LENGTH.ClientID & ".value) / 1000000,3);")

                If bnRow.Item("mFlag").ToString = "N" Then

                    mBN_CODE.CssClass = "REQUIRED"

                    If bnRow.Item("remainYN").ToString <> "Y" Then
                        If bnRow.Item("newPK").ToString = "Y" Then
                            mBN_CODE.Text = ""
                        Else
                            mBN_CODE.Text = bnRow.Item("BN_CODE").ToString
                        End If
                    End If

                    BN_X.Text = cU.FormatIntegerString(bnRow.Item("BN_X").ToString)
                    BN_Y.Text = cU.FormatIntegerString(bnRow.Item("BN_Y").ToString)
                    BN_LENGTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_LENGTH").ToString)
                    BN_WIDTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_WIDTH").ToString)
                    BN_DEPTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_DEPTH").ToString)
                    BN_REM.Text = bnRow.Item("BN_REM").ToString

                    BN_UTILIZATION_TYPE.SelectedValue = gU.decodeNullOrEmpty(bnRow.Item("BN_UTILIZATION_TYPE").ToString, "CBM")

                    BN_CBM.Text = cU.FormatDecimalwString(bnRow.Item("BN_CBM").ToString, "###,###,##0.####")
		    BN_CSMS_CODE.text = bnRow.Item("BN_CSMS_CODE").ToString
                    BN_STATUS.SelectedValue = bnRow.Item("BN_STATUS").ToString

                    sys_cb.Text = bnRow.Item("sys_cb").ToString
                    sys_lub.Text = bnRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(bnRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(bnRow.Item("sys_lud").ToString)

                    mBN_CODE.BorderWidth = Nothing
                    mBN_CODE.BackColor = Drawing.Color.Empty
                    mBN_CODE.ReadOnly = False
                Else
                    mBN_CODE.Text = bnRow.Item("BN_CODE").ToString
                    BN_X.Text = cU.FormatIntegerString(bnRow.Item("BN_X").ToString)
                    BN_Y.Text = cU.FormatIntegerString(bnRow.Item("BN_Y").ToString)
                    BN_LENGTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_LENGTH").ToString)
                    BN_WIDTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_WIDTH").ToString)
                    BN_DEPTH.Text = cU.FormatDecimalwString(bnRow.Item("BN_DEPTH").ToString)
                    BN_REM.Text = bnRow.Item("BN_REM").ToString

                    sys_cb.Text = bnRow.Item("sys_cb").ToString
                    sys_lub.Text = bnRow.Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(bnRow.Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(bnRow.Item("sys_lud").ToString)

                    BN_UTILIZATION_TYPE.SelectedValue = gU.decodeNullOrEmpty(bnRow.Item("BN_UTILIZATION_TYPE").ToString, "CBM")

                    BN_CBM.Text = cU.FormatDecimalwString(bnRow.Item("BN_CBM").ToString, "###,###,##0.####")
		    BN_CSMS_CODE.text = bnRow.Item("BN_CSMS_CODE").ToString
                    BN_STATUS.SelectedValue = bnRow.Item("BN_STATUS").ToString

                    mBN_CODE.BorderWidth = 0
                    mBN_CODE.BackColor = Drawing.Color.Transparent
                    mBN_CODE.ReadOnly = True
                End If
            Case Else
                MainTbl.Visible = False

                If Session("gLang") = "E" Then
                    lheader.Text = ""
                ElseIf Session("gLang") = "C" Then
                    lheader.Text = ""
                End If
        End Select

    End Sub

    Private Sub save()
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        Dim dupSQL As String
        Dim dupTbl As New DataTable

        Dim tmpImpCode As String
        Dim tmpFInfo As IO.FileInfo

        If Session("IMP_CODE").ToString = "" Then tmpImpCode = "W" Else tmpImpCode = Session("IMP_CODE").ToString

        Dim old_filename_FL As New ArrayList
        Dim toSave_FileName_FL As New ArrayList

        Dim old_filename_AR As New ArrayList
        Dim toSave_FileName_AR As New ArrayList

        Dim SYSP_TEMP_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")
        Dim PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR") & "\MAST_WM\"
        If Not IO.Directory.Exists(SYSP_TEMP_DIR) Then IO.Directory.CreateDirectory(SYSP_TEMP_DIR)

        Dim test2 As Integer

        test2 = arDt.Rows.Count

        'fillingUpTables("SAVE")
        test2 = arDt.Rows.Count

        generateDict()

        If StoreData(FL_NUM.Value, AR_CODE.Value, RK_CODE.Value, BN_CODE.Value) Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction

            Try
                For i As Integer = 0 To flDt.Rows.Count - 1
                    Dim fl_fileName As String = ""
                    Dim tempRemoveFLFile As Boolean = False

                    If flDt.Rows(i).Item("removeFile").ToString = "Y" Then
                        If flDt.Rows(i).Item("FL_PICTURE").ToString <> "" Then
                            old_filename_FL.Add(flDt.Rows(i).Item("FL_PICTURE").ToString)
                            tempRemoveFLFile = True
                        End If

                    ElseIf flDt.Rows(i).Item("hasFile").ToString = "Y" Then
                        If flPhoto.ContainsKey(flDt.Rows(i).Item("FL_NUM").ToString) Then
                            Dim tempPostedFile As HttpPostedFile = flPhoto.Item(flDt.Rows(i).Item("FL_NUM").ToString)

                            tmpFInfo = New FileInfo(tempPostedFile.FileName)

                            fl_fileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WHFL_" & gU.decodeNullOrEmpty(flDt.Rows(i).Item("WH_CODE").ToString, "0") & gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_NUM").ToString, "0") & tmpFInfo.Extension

                            tempPostedFile.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & fl_fileName)

                            toSave_FileName_FL.Add(fl_fileName)

                            If flDt.Rows(i).Item("FL_PICTURE").ToString <> "" Then
                                old_filename_FL.Add(flDt.Rows(i).Item("FL_PICTURE").ToString)
                            End If
                        End If
                    End If

                    If flDt.Rows(i).Item("mFlag").ToString = "N" Then

                        sql_string = "insert into wms_wh_fl ( " &
                                    "imp_code, wh_code, fl_num, " &
                                    "fl_name, fl_name_ch, fl_gross_area, " &
                                    "fl_gross_cbm, fl_max_weight, fl_length, " &
                                    "fl_width, fl_height, fl_picture, " &
                                    "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                                    gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("IMP_CODE").ToString)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("WH_CODE").ToString)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("FL_NUM").ToString)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("FL_NAME").ToString)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("FL_NAME_CH").ToString)) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_GROSS_AREA").ToString, "0")) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_GROSS_CBM").ToString, "0")) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_MAX_WEIGHT").ToString, "0")) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_LENGTH").ToString, "0")) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_WIDTH").ToString, "0")) & "," &
                                    gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_HEIGHT").ToString, "0")) & "," &
                                    gU.convdbNVCData(gU.dbEncode(fl_fileName)) & "," &
                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        sql_string = "update wms_wh_fl set " & _
                                        "fl_name = " & gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("FL_NAME").ToString)) & ", " & _
                                        "fl_name_ch = " & gU.convdbNVCData(gU.dbEncode(flDt.Rows(i).Item("FL_NAME_CH").ToString)) & ", " & _
                                        "fl_gross_area = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_GROSS_AREA").ToString, "0")) & ", " & _
                                        "fl_gross_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_GROSS_CBM").ToString, "0")) & ", " & _
                                        "fl_max_weight = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_MAX_WEIGHT").ToString, "0")) & ", " & _
                                        "fl_length = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_LENGTH").ToString, "0")) & ", " & _
                                        "fl_width = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_WIDTH").ToString, "0")) & ", " & _
                                        "fl_height = " & gU.dbEncode(gU.decodeNullOrEmpty(flDt.Rows(i).Item("FL_HEIGHT").ToString, "0")) & ", "

                        If Not tempRemoveFLFile Then
                            If fl_fileName <> "" Then sql_string = sql_string & "fl_picture = " & gU.convdbNVCData(gU.dbEncode(gU.dbEncode(fl_fileName))) & ", "
                        Else
                            sql_string = sql_string & "fl_picture = null, "
                        End If

                        sql_string = sql_string & "sys_lub = '" & Session("usr_id") & "', " & _
                                        "sys_lud = Getdate() " & _
                                        "where imp_code = '" & gU.dbEncode(flDt.Rows(i).Item("IMP_CODE").ToString) & "' " & _
                                        "and wh_code = '" & gU.dbEncode(flDt.Rows(i).Item("WH_CODE").ToString) & "' " & _
                                        "and fl_num = '" & gU.dbEncode(flDt.Rows(i).Item("FL_NUM").ToString) & "' "
                    End If

                    If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)
                Next

                Dim arRow As DataRow
                For Each arkeyPair As KeyValuePair(Of String, DataRow) In arDict
                    arRow = arkeyPair.Value

                    Dim ar_fileName As String = ""
                    Dim tempRemoveARFile As Boolean = False

                    If arRow.Item("removeFile").ToString = "Y" Then
                        If arRow.Item("AR_PICTURE").ToString <> "" Then
                            old_filename_AR.Add(arRow.Item("ar_PICTURE").ToString)
                            tempRemoveARFile = True
                        End If

                    ElseIf arRow.Item("hasFile").ToString = "Y" Then
                        If arPhoto.ContainsKey(arRow.Item("AR_CODE").ToString) Then
                            Dim tempPostedFile As HttpPostedFile = arPhoto.Item(arRow.Item("AR_CODE").ToString)

                            tmpFInfo = New FileInfo(tempPostedFile.FileName)

                            ar_fileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WHAR_" & gU.decodeNullOrEmpty(arRow.Item("WH_CODE").ToString, "0") & gU.decodeNullOrEmpty(arRow.Item("AR_CODE").ToString, "0") & tmpFInfo.Extension

                            tempPostedFile.SaveAs(SYSP_TEMP_DIR & "\" & "temp_" & ar_fileName)

                            toSave_FileName_AR.Add(ar_fileName)

                            If arRow.Item("AR_PICTURE").ToString <> "" Then
                                old_filename_AR.Add(arRow.Item("AR_PICTURE").ToString)
                            End If
                        End If
                    End If

                    If arRow.Item("mFlag").ToString = "N" Then
                        sql_string = "insert into WMS_WH_AREA ( " &
                                   "imp_code, wh_code, fl_num, " &
                                   "ar_code, ar_name, ar_name_ch, " &
                                   "ar_x, ar_y, ar_length, " &
                                   "ar_width, ar_height, ar_gross_area, " &
                                   "ar_gross_cbm, ar_facilities, ar_picture, ar_damage_yn, ar_temp_fr, ar_temp_to, ar_temp_yn,ar_power,ar_security, " &
                                   "AR_SCRAP_AREA, AR_SHORTLEN_CABLE_AREA, AR_INSP_AREA, AR_REPAIR_AREA, AR_PICKDROP_AREA, AR_FFI, " &
                                   "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("IMP_CODE").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("WH_CODE").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("FL_NUM").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_CODE").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_NAME").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_NAME_CH").ToString)) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_X").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_Y").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_LENGTH").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_WIDTH").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_HEIGHT").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_GROSS_AREA").ToString, "0")) & "," &
                                   gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_GROSS_CBM").ToString, "0")) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_FACILITIES").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(ar_fileName)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_damage_yn").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_temp_fr").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_temp_to").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_temp_yn").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_power").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_security").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_SCRAP_AREA").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_SHORTLEN_CABLE_AREA").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_INSP_AREA").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_REPAIR_AREA").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_PICKDROP_AREA").ToString)) & "," &
                                   gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_FFI").ToString)) & "," &
                                   "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        sql_string = "update WMS_WH_AREA set " & _
                                  "AR_NAME = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_NAME").ToString)) & ", " & _
                                  "AR_NAME_CH = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_NAME_CH").ToString)) & ", " & _
                                  "AR_X = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_X").ToString, "0")) & ", " & _
                                  "AR_Y = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_Y").ToString, "0")) & ", " & _
                                  "AR_LENGTH = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_LENGTH").ToString, "0")) & ", " & _
                                  "AR_WIDTH = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_WIDTH").ToString, "0")) & ", " & _
                                  "AR_HEIGHT = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_HEIGHT").ToString, "0")) & ", " & _
                                  "AR_GROSS_AREA = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_GROSS_AREA").ToString, "0")) & ", " & _
                                  "AR_GROSS_CBM = " & gU.dbEncode(gU.decodeNullOrEmpty(arRow.Item("AR_GROSS_CBM").ToString, "0")) & ", " & _
                                  "AR_FACILITIES = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_FACILITIES").ToString)) & ", " & _
                                  "AR_DAMAGE_YN = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_DAMAGE_YN").ToString)) & ", " & _
                                  "AR_TEMP_FR = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_TEMP_FR").ToString)) & ", " & _
                                  "AR_TEMP_TO = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_TEMP_TO").ToString)) & ", " & _
                                  "ar_power = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_power").ToString)) & ", " & _
                                  "ar_security = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_security").ToString)) & ", " & _
                                  "ar_temp_yn = " & gU.convdbNVCData(gU.dbEncode(arRow.Item("ar_temp_yn").ToString)) & ", " & _
                                  "AR_SCRAP_AREA=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_SCRAP_AREA").ToString)) & ", " & _
                                  "AR_SHORTLEN_CABLE_AREA=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_SHORTLEN_CABLE_AREA").ToString)) & ", " & _
                                  "AR_INSP_AREA=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_INSP_AREA").ToString)) & ", " & _
                                  "AR_REPAIR_AREA=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_REPAIR_AREA").ToString)) & ", " & _
                                  "AR_PICKDROP_AREA=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_PICKDROP_AREA").ToString)) & ", " & _
                                  "AR_FFI=" & gU.convdbNVCData(gU.dbEncode(arRow.Item("AR_FFI").ToString)) & ", "

                        If Not tempRemoveARFile Then
                            If ar_fileName <> "" Then sql_string = sql_string & "ar_picture = " & gU.convdbNVCData(gU.dbEncode(gU.dbEncode(ar_fileName))) & ", "
                        Else
                            sql_string = sql_string & "ar_picture = null, "
                        End If

                        sql_string = sql_string & "sys_lub = '" & Session("usr_id") & "', " & _
                                  "sys_lud = Getdate() " & _
                                  "where imp_code = '" & gU.dbEncode(arRow.Item("IMP_CODE").ToString) & "' " & _
                                  "and WH_CODE = '" & gU.dbEncode(arRow.Item("WH_CODE").ToString) & "' " & _
                                  "and FL_NUM = '" & gU.dbEncode(arRow.Item("FL_NUM").ToString) & "' " & _
                                  "and AR_CODE = '" & gU.dbEncode(arRow.Item("AR_CODE").ToString) & "' "
                    End If

                    If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)
                Next

                Dim rkRow As DataRow
                For Each rkkeyPair As KeyValuePair(Of String, DataRow) In rkDict
                    rkRow = rkkeyPair.Value

                    If rkRow.Item("mFlag").ToString = "N" Then
                        sql_string = "insert into WMS_WH_RACK ( " &
                                 "imp_code, wh_code, fl_num, " &
                                 "ar_code, rk_code, RK_NAME, " &
                                 "RK_NAME_CH, RK_GROSS_AREA, RK_GROSS_CBM, " &
                                 "RK_X, RK_Y, RK_LENGTH, " &
                                 "RK_WIDTH, RK_DEPTH, RK_XBINS, " &
                                 "RK_YBINS, RK_REM, " &
                                 "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("IMP_CODE").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("WH_CODE").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("FL_NUM").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("AR_CODE").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_CODE").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_NAME").ToString)) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_NAME_CH").ToString)) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_GROSS_AREA").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_GROSS_CBM").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_X").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_Y").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_LENGTH").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_WIDTH").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_DEPTH").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_XBINS").ToString, "0")) & "," &
                                 gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_YBINS").ToString, "0")) & "," &
                                 gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_REM").ToString)) & "," &
                                 "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        sql_string = "update WMS_WH_RACK set " & _
                                  "RK_NAME = " & gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_NAME").ToString)) & ", " & _
                                  "RK_NAME_CH = " & gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_NAME_CH").ToString)) & ", " & _
                                  "RK_GROSS_AREA = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_GROSS_AREA").ToString, "0")) & ", " & _
                                  "RK_GROSS_CBM = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_GROSS_CBM").ToString, "0")) & ", " & _
                                  "RK_X = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_X").ToString, "0")) & ", " & _
                                  "RK_Y = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_Y").ToString, "0")) & ", " & _
                                  "RK_LENGTH = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_LENGTH").ToString, "0")) & ", " & _
                                  "RK_WIDTH = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_WIDTH").ToString, "0")) & ", " & _
                                  "RK_DEPTH = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_DEPTH").ToString, "0")) & ", " & _
                                  "RK_XBINS = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_XBINS").ToString, "0")) & ", " & _
                                  "RK_YBINS = " & gU.dbEncode(gU.decodeNullOrEmpty(rkRow.Item("RK_YBINS").ToString, "0")) & ", " & _
                                  "RK_REM = " & gU.convdbNVCData(gU.dbEncode(rkRow.Item("RK_REM").ToString)) & ", " & _
                                  "sys_lub = '" & Session("usr_id") & "', " & _
                                  "sys_lud = Getdate() " & _
                                  "where imp_code = '" & gU.dbEncode(rkRow.Item("IMP_CODE").ToString) & "' " & _
                                  "and WH_CODE = '" & gU.dbEncode(rkRow.Item("WH_CODE").ToString) & "' " & _
                                  "and FL_NUM = '" & gU.dbEncode(rkRow.Item("FL_NUM").ToString) & "' " & _
                                  "and AR_CODE = '" & gU.dbEncode(rkRow.Item("AR_CODE").ToString) & "' " & _
                                  "and RK_CODE = '" & gU.dbEncode(rkRow.Item("RK_CODE").ToString) & "' "
                    End If

                    If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)


                    sql_string = "delete from WMS_WH_BIN " & _
                                 "where imp_code = '" & gU.dbEncode(rkRow.Item("IMP_CODE").ToString) & "' " & _
                                 "and WH_CODE = '" & gU.dbEncode(rkRow.Item("WH_CODE").ToString) & "' " & _
                                 "and FL_NUM = '" & gU.dbEncode(rkRow.Item("FL_NUM").ToString) & "' " & _
                                 "and AR_CODE = '" & gU.dbEncode(rkRow.Item("AR_CODE").ToString) & "' " & _
                                 "and RK_CODE = '" & gU.dbEncode(rkRow.Item("RK_CODE").ToString) & "' " & _
                                 "and not exists ( " & _
                                    "select 1 from WMS_WH_RACK r " & _
                                    "where r.imp_code = WMS_WH_BIN.imp_code " & _
                                    "and r.wh_code = WMS_WH_BIN.wh_code " & _
                                    "and r.fl_num = WMS_WH_BIN.fl_num " & _
                                    "and r.ar_code = WMS_WH_BIN.ar_code " & _
                                    "and r.rk_code = WMS_WH_BIN.rk_code " & _
                                    "and (WMS_WH_BIN.bn_x <= r.rk_x and WMS_WH_BIN.bn_y <= r.rk_y)) "

                    gDB.amendData(sql_string, gConn, transaction)
                Next


                Dim bnRow As DataRow
                For Each bnkeyPair As KeyValuePair(Of String, DataRow) In bnDict
                    bnRow = bnkeyPair.Value

                    If bnRow.Item("mFlag").ToString = "N" Then

                        If bnRow.Item("BN_CODE").ToString = ViewState("pre_bncode").ToString Then
                            If bnRow.Item("mFlag") = "N" And bnRow.Item("newPK") = "Y" And mBN_CODE.Text = "" Then
                                Continue For
                            End If
                        End If


                        sql_string = "insert into WMS_WH_BIN ( " &
                                         "imp_code, wh_code, fl_num, " &
                                         "ar_code, rk_code, bn_code, " &
                                         "BN_X, BN_Y, BN_LENGTH, " &
                                         "BN_WIDTH, BN_DEPTH, BN_REM, BN_UTILIZATION_TYPE,BN_STATUS,BN_CBM,BN_CSMS_CODE," &
                                         "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("IMP_CODE").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("WH_CODE").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("FL_NUM").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("AR_CODE").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("RK_CODE").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_CODE").ToString)) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_X").ToString, "0")) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_Y").ToString, "0")) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_LENGTH").ToString, "0")) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_WIDTH").ToString, "0")) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_DEPTH").ToString, "0")) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_REM").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_UTILIZATION_TYPE").ToString)) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_STATUS").ToString)) & "," &
                                         gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_CBM").ToString, "0")) & "," &
                                         gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_CSMS_CODE").ToString)) & "," &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        sql_string = "update WMS_WH_BIN set " & _
                                "BN_X = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_X").ToString, "0")) & ", " & _
                                "BN_Y = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_Y").ToString, "0")) & ", " & _
                                "BN_LENGTH = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_LENGTH").ToString, "0")) & ", " & _
                                "BN_WIDTH = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_WIDTH").ToString, "0")) & ", " & _
                                "BN_DEPTH = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_DEPTH").ToString, "0")) & ", " & _
                                "BN_REM = " & gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_REM").ToString)) & ", " & _
                                "BN_UTILIZATION_TYPE = " & gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_UTILIZATION_TYPE").ToString)) & ", " & _
                                "BN_STATUS = " & gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_STATUS").ToString)) & ", " & _
                                "BN_CBM = " & gU.dbEncode(gU.decodeNullOrEmpty(bnRow.Item("BN_CBM").ToString, "0")) & ", " & _
                                "BN_CSMS_CODE = " & gU.convdbNVCData(gU.dbEncode(bnRow.Item("BN_CSMS_CODE").ToString)) & ", " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(bnRow.Item("IMP_CODE").ToString) & "' " & _
                                "and WH_CODE = '" & gU.dbEncode(bnRow.Item("WH_CODE").ToString) & "' " & _
                                "and FL_NUM = '" & gU.dbEncode(bnRow.Item("FL_NUM").ToString) & "' " & _
                                "and AR_CODE = '" & gU.dbEncode(bnRow.Item("AR_CODE").ToString) & "' " & _
                                "and RK_CODE = '" & gU.dbEncode(bnRow.Item("RK_CODE").ToString) & "' " & _
                                "and BN_CODE = '" & gU.dbEncode(bnRow.Item("BN_CODE").ToString) & "' "
                                

                    End If

                    If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)
                Next

                If Not transaction Is Nothing Then
                    transaction.Commit()
                    transaction = Nothing

                    For x As Integer = 0 To old_filename_FL.Count - 1
                        If IO.File.Exists(PHY_PS_DIR & old_filename_FL(x)) Then
                            thumb = New ThumbGenerator
                            thumb.SetParams(PHY_PS_DIR & old_filename_FL(x), 150, 180)
                            Dim unID As String = thumb.GetUniqueThumbName
                            Cache.Remove(unID)

                            IO.File.Delete(PHY_PS_DIR & old_filename_FL(x))
                        End If
                    Next

                    For x As Integer = 0 To toSave_FileName_FL.Count - 1
                        If IO.File.Exists(SYSP_TEMP_DIR & "\" & "temp_" & toSave_FileName_FL(x)) Then
                            IO.File.Move(SYSP_TEMP_DIR & "\" & "temp_" & toSave_FileName_FL(x), PHY_PS_DIR & toSave_FileName_FL(x))
                        End If
                    Next

                    For i As Integer = 0 To old_filename_AR.Count - 1
                        If IO.File.Exists(PHY_PS_DIR & old_filename_AR(i)) Then
                            thumb = New ThumbGenerator
                            thumb.SetParams(PHY_PS_DIR & old_filename_AR(i), 150, 180)
                            Dim unID As String = thumb.GetUniqueThumbName
                            Cache.Remove(unID)

                            IO.File.Delete(PHY_PS_DIR & old_filename_AR(i))
                        End If
                    Next

                    For i As Integer = 0 To toSave_FileName_AR.Count - 1
                        If IO.File.Exists(SYSP_TEMP_DIR & "\" & "temp_" & toSave_FileName_AR(i)) Then
                            IO.File.Move(SYSP_TEMP_DIR & "\" & "temp_" & toSave_FileName_AR(i), PHY_PS_DIR & toSave_FileName_AR(i))
                        End If
                    Next
                End If

                uiFun.displayMsg(Me, "1001", "", Session("gLang"))

                'Session("arDict") = Nothing
                'Session("rkDict") = Nothing
                'Session("bnDict") = Nothing
                'Session("flPhoto") = Nothing

                arDict.Clear()
                rkDict.Clear()
                bnDict.Clear()
                flPhoto.Clear()

                'FL_PICTURE_upload = New FileUpload

                ViewState("first_load") = True
                ViewState("loadbn") = True
                ViewState("flDt") = Nothing
                ViewState("arDt") = Nothing
                ViewState("rkDt") = Nothing
                ViewState("bnDt") = Nothing
                ViewState("bnsaveDt") = Nothing

                mode.Value = ""
                GenerateAll()

            Catch ex As Exception
                transaction.Rollback()
                uiFun.displayMsg(Me, "1002", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        End If
    End Sub

    Private Sub AddFloor()

        Dim canContinue As Boolean = True

        If flTr.Visible = True Then
            If mFL_NUM.Text.Trim = "" Or FL_NAME.Text = "" Then
                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf arTr.Visible = True Then
            If mAR_CODE.Text.Trim = "" Or AR_NAME.Text = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf rkTr.Visible = True Then
            If mRK_CODE.Text.Trim = "" Or RK_NAME.Text = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf ViewState("notAllowAdd") = True Then

            mode.Value = ""

            GenerateAll()

            canContinue = False
        End If

        If canContinue Then
            Dim rows_count As Integer = 0
            Dim tempRow As DataRow = flDt.NewRow

            Dim ran As New Random
            Dim ranNum As Integer
            Dim dCount As Integer = 1

            Do While dCount > 0
                ranNum = CInt(Now.Millisecond) + CInt(Now.Minute) + ran.Next(1, 9999)
                dCount = flDt.Compute("COUNT(FL_NUM)", "WH_CODE = '" & WH_CODE.Value & "' AND FL_NUM = '" & ranNum & "'")
            Loop

            rows_count = flDt.Rows.Count

            tempRow.Item("IMP_CODE") = Session("IMP_CODE").ToString
            tempRow.Item("WH_CODE") = WH_CODE.Value
            tempRow.Item("FL_NUM") = ranNum
            tempRow.Item("mFlag") = "N"
            tempRow.Item("newPK") = "Y"
            tempRow.Item("FL_NAME_D") = "NEW"

            FL_NUM.Value = tempRow.Item("FL_NUM")

            flDt.Rows.Add(tempRow)

            flDt.AcceptChanges()

            mode.Value = ""

            GenerateAll()
        End If
    End Sub

    Private Sub AddArea()
        Dim rows_count As Integer = 0
        Dim tempRow As DataRow
        Dim ranNum As New Random

        Dim canContinue As Boolean = True

        If flTr.Visible = True Then
            If mFL_NUM.Text.Trim = "" Or FL_NAME.Text.Trim = "" Then
                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf arTr.Visible = True Then
            If mAR_CODE.Text.Trim = "" Or AR_NAME.Text.Trim = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf rkTr.Visible = True Then
            If mRK_CODE.Text.Trim = "" Or RK_NAME.Text.Trim = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf ViewState("notAllowAdd") = True Then

            mode.Value = ""

            GenerateAll()

            canContinue = False
        End If

        If canContinue Then
            rows_count = arDict.Count

            If FL_NUM.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Floor First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料重複!!", Session("gLang"))
                End If
            Else

                Dim arSchema As New DataTable
                Dim getSql As String = "select top 1 *, AR_NAME as AR_NAME_D, '' as mFlag, '' as newPK, '' as remainYN, '' as hasFile, '' as removeFile from WMS_WH_AREA"
                Dim fNum As String = ""

                If flTr.Visible = True Then
                    fNum = mFL_NUM.Text.Trim
                    FL_NUM.Value = fNum
                Else
                    fNum = FL_NUM.Value
                End If

                gDB.getDataTable(getSql, , , arSchema)

                If arSchema.Columns.Count > 0 Then
                    'tempRow = arSchema.NewRow
                    tempRow = arDt.NewRow

                    tempRow.Item("IMP_CODE") = Session("IMP_CODE").ToString
                    tempRow.Item("WH_CODE") = WH_CODE.Value
                    tempRow.Item("FL_NUM") = fNum
                    tempRow.Item("AR_CODE") = rows_count + 1 + CInt(Now.Millisecond) + CInt(Now.Minute) + ranNum.Next(1, 2999)
                    tempRow.Item("mFlag") = "N"
                    tempRow.Item("newPK") = "Y"
                    tempRow.Item("AR_NAME_D") = "NEW"

                    AR_CODE.Value = tempRow.Item("AR_CODE")

                    'arDict.Add((fNum & "||" & AR_CODE.Value), tempRow)

                    arDt.Rows.Add(tempRow)

                    arDt.AcceptChanges()

                    'Session("arDict") = arDict

                    arSchema.Dispose()

                    mode.Value = ""

                    RK_CODE.Value = ""

                    GenerateAll("ADD")

                End If
            End If
        End If
    End Sub

    Private Sub AddRack()
        Dim rows_count As Integer = 0
        Dim tempRow As DataRow
        Dim ranNum As New Random

        Dim canContinue As Boolean = True

        If flTr.Visible = True Then
            If mFL_NUM.Text.Trim = "" Or FL_NAME.Text.Trim = "" Then
                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf arTr.Visible = True Then
            If mAR_CODE.Text.Trim = "" Or AR_NAME.Text.Trim = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf rkTr.Visible = True Then
            If mRK_CODE.Text.Trim = "" Or RK_NAME.Text.Trim = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf ViewState("notAllowAdd") = True Then

            mode.Value = ""

            GenerateAll()

            canContinue = False
        End If

        If canContinue Then
            rows_count = rkDict.Count

            If FL_NUM.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Floor First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料!!", Session("gLang"))
                End If
            ElseIf AR_CODE.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Area First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料!!", Session("gLang"))
                End If
            Else
                Dim rkSchema As New DataTable
                Dim getSql As String = "select top 1 *, RK_NAME as RK_NAME_D, '' as mFlag, '' as newPK, '' as remainYN from WMS_WH_RACK"

                Dim arC As String = ""

                If arTr.Visible = True Then
                    arC = mAR_CODE.Text.Trim
                    AR_CODE.Value = arC
                Else
                    arC = AR_CODE.Value
                End If

                gDB.getDataTable(getSql, , , rkSchema)

                If rkSchema.Columns.Count > 0 Then
                    'tempRow = rkSchema.NewRow
                    tempRow = rkDt.NewRow

                    tempRow.Item("IMP_CODE") = Session("IMP_CODE").ToString
                    tempRow.Item("WH_CODE") = WH_CODE.Value
                    tempRow.Item("FL_NUM") = FL_NUM.Value
                    tempRow.Item("AR_CODE") = arC
                    tempRow.Item("RK_CODE") = rows_count + 1 + CInt(Now.Millisecond) + CInt(Now.Minute) + ranNum.Next(1, 2999)
                    tempRow.Item("mFlag") = "N"
                    tempRow.Item("newPK") = "Y"
                    tempRow.Item("RK_NAME_D") = "NEW"

                    RK_CODE.Value = tempRow.Item("RK_CODE")

                    '                    rkDict.Add((FL_NUM.Value & "||" & arC & "||" & RK_CODE.Value), tempRow)
                    rkDt.Rows.Add(tempRow)

                    rkDt.AcceptChanges()

                    'Session("rkDict") = rkDict

                    rkSchema.Dispose()

                    mode.Value = ""

                    GenerateAll("ADD")
                End If
            End If
        End If
    End Sub

    Private Sub AddBin()
        Dim canContinue As Boolean = True

        If rkTr.Visible = True Then
            If mRK_CODE.Text.Trim = "" Or RK_NAME.Text.Trim = "" Then

                mode.Value = ""

                GenerateAll()

                canContinue = False
            End If
        ElseIf bnTr.Visible = True Then
            If mBN_CODE.Text.Trim = "" Then

                'mode.Value = ""

                'GenerateAll()

                'canContinue = False
            End If
        ElseIf ViewState("notAllowAdd") = True Then

            mode.Value = ""

            GenerateAll()

            canContinue = False
        End If

        If canContinue Then

            bnDt = ViewState("bnDt")

            Dim rows_count As Integer = 0
            Dim tempRow As DataRow
            Dim ranNum As New Random

            rows_count = bnDict.Count

            If FL_NUM.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Floor First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料!!", Session("gLang"))
                End If
            ElseIf AR_CODE.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Area First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料!!", Session("gLang"))
                End If
            ElseIf RK_CODE.Value = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select Rack First!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "資料!!", Session("gLang"))
                End If
            Else
                Dim bnSchema As New DataTable
                Dim getSql As String = "select top 1 *, BN_CODE as BN_CODE_D, '' as mFlag, '' as newPK, '' as remainYN from WMS_WH_BIN "

                Dim rkC As String = ""

                If rkTr.Visible = True Then
                    rkC = mRK_CODE.Text
                    RK_CODE.Value = rkC
                Else
                    rkC = RK_CODE.Value
                End If

                gDB.getDataTable(getSql, , , bnSchema)

                If bnSchema.Columns.Count > 0 Then
                    Dim y_asix As String = ""
                    Dim x_asix As String = ""

                    Dim tempXY As String()

                    If BN_CODE.Value.Contains("_E_") Then
                        tempXY = Split(BN_CODE.Value, "_E_")

                        y_asix = tempXY(0)
                        x_asix = tempXY(1)

                        'tempRow = bnSchema.NewRow
                        tempRow = bnsaveDt.NewRow

                        tempRow.Item("IMP_CODE") = Session("IMP_CODE").ToString
                        tempRow.Item("WH_CODE") = WH_CODE.Value
                        tempRow.Item("FL_NUM") = FL_NUM.Value
                        tempRow.Item("AR_CODE") = AR_CODE.Value
                        tempRow.Item("RK_CODE") = rkC
                        tempRow.Item("BN_CODE") = rows_count + 1 + CInt(Now.Month) + CInt(Now.Day) + CInt(Now.Year) + CInt(Now.Millisecond) + CInt(Now.Minute) + ranNum.Next(1, 2999)
                        tempRow.Item("BN_X") = (CInt(x_asix) + 1).ToString
                        tempRow.Item("BN_Y") = (CInt(y_asix) + 1).ToString
                        tempRow.Item("mFlag") = "N"
                        tempRow.Item("newPK") = "Y"
                        tempRow.Item("BN_CODE_D") = "NEW"
                        tempRow.Item("BN_STATUS") = "ACTIVE"

                        BN_CODE.Value = tempRow.Item("BN_CODE")

                        bnDt.Rows(CInt(y_asix)).Item(CInt(x_asix)) = BN_CODE.Value

                        'bnDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkC & "||" & BN_CODE.Value), tempRow)
                        bnsaveDt.Rows.Add(tempRow)

                        bnsaveDt.AcceptChanges()

                        'Session("bnDict") = bnDict

                        bnSchema.Dispose()

                        mode.Value = ""

                        generateDict()

                        GenerateAll("ADD")
                    End If


                End If
            End If
        End If
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        save()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        save()
    End Sub

    Private Sub removeImage(ByVal FileName As String)
        If FileName <> "" Then
            Try
                Dim PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR") & "\MAST_WM\"

                If IO.File.Exists(PHY_PS_DIR & FileName) Then

                    thumb = New ThumbGenerator
                    thumb.SetParams(PHY_PS_DIR & FileName, 150, 180)
                    Dim unID As String = thumb.GetUniqueThumbName
                    Cache.Remove(unID)

                    IO.File.Delete(PHY_PS_DIR & FileName)
                End If

            Catch ex As Exception
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Private Function validateAll() As Boolean
        Return True
    End Function

    Protected Sub FL_PICTURE_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FL_PICTURE_remove.Click
        Dim nDataRows() As DataRow = flDt.Select("FL_NUM = '" & FL_NUM.Value & "'")
        If nDataRows.Count > 0 Then
            nDataRows(0).Item("removeFile") = "Y"

            FL_PICTURE_remove.Enabled = False
            FL_PICTURE.Enabled = False
            fl_preview_pic.Enabled = False
            FL_PICTURE_upload.Visible = True
            FL_PICTURE_upload.Enabled = False
            FL_PICTURE_edit.Visible = False
            FL_PICTURE_lit.Text = ""
            FL_PICTURE_lit.Visible = False
        End If
    End Sub

    Protected Sub FL_PICTURE_edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FL_PICTURE_edit.Click
        FL_PICTURE_upload.Visible = True
        FL_PICTURE_upload.Enabled = True
        FL_PICTURE_lit.Text = ""
        FL_PICTURE_lit.Visible = False
        FL_PICTURE_edit.Visible = False
    End Sub

    Protected Sub AR_PICTURE_REMOVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AR_PICTURE_REMOVE.Click
        Dim dicCode As String = FL_NUM.Value & "||" & AR_CODE.Value
        Dim arRow As DataRow = arDict.Item(dicCode)

        If arRow IsNot Nothing Then
            arRow.Item("removeFile") = "Y"

            AR_PICTURE_REMOVE.Enabled = False
            AR_PICTURE.Enabled = False
            ar_preview_pic.Enabled = False
            AR_PICTURE_UPLOAD.Visible = True
            AR_PICTURE_UPLOAD.Enabled = False
            AR_PICTURE_EDIT.Visible = False
            AR_PICTURE_LIT.Text = ""
            AR_PICTURE_LIT.Visible = False
        End If
    End Sub

    Protected Sub AR_PICTURE_EDIT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AR_PICTURE_EDIT.Click
        AR_PICTURE_UPLOAD.Visible = True
        AR_PICTURE_UPLOAD.Enabled = True
        AR_PICTURE_LIT.Text = ""
        AR_PICTURE_LIT.Visible = False
        AR_PICTURE_EDIT.Visible = False
    End Sub

    Private Function generateDict() As Boolean

        For i As Integer = 0 To arDt.Rows.Count - 1
            If Not arDict.ContainsKey((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString)) Then
                arDict.Add((FL_NUM.Value & "||" & arDt.Rows(i).Item("AR_CODE").ToString), arDt.Rows(i))
            End If
        Next

        For i As Integer = 0 To rkDt.Rows.Count - 1
            If Not rkDict.ContainsKey((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString)) Then
                rkDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & rkDt.Rows(i).Item("RK_CODE").ToString), rkDt.Rows(i))
            End If
        Next

        For i As Integer = 0 To bnsaveDt.Rows.Count - 1
            If Not bnDict.ContainsKey((FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnsaveDt.Rows(i).Item("BN_CODE").ToString)) Then
                bnDict.Add((FL_NUM.Value & "||" & AR_CODE.Value & "||" & RK_CODE.Value & "||" & bnsaveDt.Rows(i).Item("BN_CODE").ToString), bnsaveDt.Rows(i))
            End If
        Next
        Return True
    End Function

End Class