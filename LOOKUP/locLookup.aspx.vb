Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class locLookup
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private pItemArray As String()
    Private newDataArray As String()
    Private lWH_CODE As String = ""
    Private lFL_NUM As String = ""
    Private paP As GlobalDBFunc.DBCmdPara


    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request("pForm") <> "" Then
            pForm.value = Request("pForm")
        End If
        If Request("pItemList") <> "" Then
            pItemList.value = Request("pItemList")
        End If

        pItemArray = Split(pItemList.value, ", ")
        newDataArray = Split(pItemList.Value, ", ")

        If Request("mwh") <> "" Then
            mwh.Value = Request("mwh")
            wh.Value = ""
        Else
            If Request("wh") <> "" Then
                wh.Value = Request("wh")
            End If
        End If

        
        If Request("sc") <> "" Then
            sc.Value = Request("sc")
        End If
        If Request("ic") <> "" Then
            itemCode.Value = Request("ic")
        End If
        If Request("pack") <> "" Then
            packkey.Value = Request("pack")
        End If

        If Request("fun_code") <> "" Then
            FUN_CODE.Value = Request("fun_code")
        End If

        'Session("IMP_CODE") = "WMS"

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lblTitle.Text = "Location Lookup"
            'lblRK_CODE.Text = "Rack Code"
            'lblBN_CODE.text = "Bn Code"
            Submit.Text = "Select"
        Else
            lblTitle.Text = "Location Lookup"
            'lblRK_CODE.Text = "Rack Code"
            'lblBN_CODE.text = "Bn Code"
            Submit.Text = "Select"
        End If
        REM **********************


        If Not IsPostBack Then
            REM**********************
            REM Generate Dropdown List from WMS_COL_CODE Table
            'uiFun.load_dropdown(WH_CODE, "select WH_CODE, WH_NAME from WMS_WAREHOUSE", "WH_CODE", "WH_NAME")
            REM **********************
        End If

        REM Generate Warehouse
        Dim colCount As Integer = 0

        Dim tclblWH_CODE As New TableCell()
        tclblWH_CODE.CssClass = "LabelTD"
        tclblWH_CODE.Style.Add("width", "20%")
        Dim WH_CODElabel As New Label
        WH_CODElabel.Text = "Subinventory"
        tclblWH_CODE.Controls.Add(WH_CODElabel)
        trWH_CODE.Cells.Add(tclblWH_CODE)

        Dim tcWH_CODE As New TableCell()
        tcWH_CODE.Style.Add("width", "80%")

        Dim srch_sql As String = ""
        srch_sql = "select WH_CODE, WH_NAME from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"

        If mwh.Value <> "" Then
            srch_sql += " AND WH_MAIN_WH = '" & gU.dbEncode(mwh.Value) & "'"
        Else
            If wh.Value <> "" Then
                srch_sql += " AND WH_CODE = '" & gU.dbEncode(wh.Value) & "'"
            End If
        End If

        srch_sql += " order by WH_NAME"

        Dim dt As DataTable
        dt = gDB.getDataTable(srch_sql)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nbutton As New Button
            nbutton.Width = 125
            nbutton.CssClass = "loc_button"
            nbutton.ID = "WH_CODE_" & dt.Rows(i).Item(0).ToString
            If WH_CODE.Value = dt.Rows(i).Item(0).ToString Then
                'nbutton.Text = dt.Rows(i).Item(1).ToString & " * "
                nbutton.Text = dt.Rows(i).Item(1).ToString
                nbutton.CssClass = "selected_button"
            Else
                nbutton.Text = dt.Rows(i).Item(1).ToString
            End If
            nbutton.OnClientClick = "window.document.forms[0].WH_CODE.value=""" & gU.jsHTMLEncode(dt.Rows(i).Item(0).ToString) & """;window.document.forms[0].FL_NUM.value="""";window.document.forms[0].AR_CODE.value="""";;window.document.forms[0].RK_CODE.value="""";"
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
        trWH_CODE.Cells.Add(tcWH_CODE)

        REM Generate FL
        colCount = 0
        Dim tclbltrFL_NUM As New TableCell()
        tclbltrFL_NUM.CssClass = "LabelTD"
        tclbltrFL_NUM.Style.Add("width", "20%")
        Dim FL_NUMlabel As New Label
        FL_NUMlabel.Text = "Floor"
        tclbltrFL_NUM.Controls.Add(FL_NUMlabel)
        trFL_NUM.Cells.Add(tclbltrFL_NUM)

        Dim tcFL_NUM As New TableCell()
        tcFL_NUM.Style.Add("width", "80%")

        srch_sql = "select FL_NUM, FL_NAME from WMS_WH_FL WHERE WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "' AND IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"

        srch_sql += " order by FL_NAME"

        dt = gDB.getDataTable(srch_sql)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nbutton As New Button
            nbutton.Width = 75
            nbutton.CssClass = "loc_button"
            nbutton.ID = "FL_NUM_" & dt.Rows(i).Item(0).ToString
            If FL_NUM.Value = dt.Rows(i).Item(0).ToString Then
                'nbutton.Text = dt.Rows(i).Item(1).ToString & " * "
                nbutton.Text = dt.Rows(i).Item(1).ToString
                nbutton.CssClass = "selected_button"
            Else
                nbutton.Text = dt.Rows(i).Item(1).ToString
            End If
            nbutton.OnClientClick = "window.document.forms[0].FL_NUM.value=""" & gU.jsHTMLEncode(dt.Rows(i).Item(0).ToString) & """;window.document.forms[0].AR_CODE.value="""";;window.document.forms[0].RK_CODE.value="""";"
            'AddHandler nbutton.Click, AddressOf WHCODEClick

            tcFL_NUM.Controls.Add(nbutton)

            colCount += 1

            If colCount = 5 Then
                Dim lit As New Literal
                lit.Text = "<br/>"
                tcWH_CODE.Controls.Add(lit)
                colCount = 0
            End If
        Next
        trFL_NUM.Cells.Add(tcFL_NUM)

        REM Generate AR
        colCount = 0
        Dim tclbltrAR_CODE As New TableCell()
        tclbltrAR_CODE.CssClass = "LabelTD"
        tclbltrAR_CODE.Style.Add("width", "20%")
        Dim AR_CODElabel As New Label
        AR_CODElabel.Text = "Area"
        tclbltrAR_CODE.Controls.Add(AR_CODElabel)
        trAR_CODE.Cells.Add(tclbltrAR_CODE)

        Dim tcAR_CODE As New TableCell()
        tcAR_CODE.Style.Add("width", "80%")

        paP = New GlobalDBFunc.DBCmdPara
        srch_sql = "select AR_CODE, AR_NAME from WMS_WH_AREA WHERE WH_CODE =" & paP.AP(WH_CODE.Value) & " and FL_NUM = '" & FL_NUM.Value & "' AND IMP_CODE = " & paP.AP(Session("IMP_CODE"))
        srch_sql += " order by AR_NAME"

        dt = gDB.getDataTable(srch_sql, , , , paP)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nbutton As New Button
            nbutton.Width = 75
            nbutton.CssClass = "loc_button"
            nbutton.ID = "AR_CODE_" & dt.Rows(i).Item(0).ToString
            If AR_CODE.Value = dt.Rows(i).Item(0).ToString Then
                'nbutton.Text = dt.Rows(i).Item(1).ToString & " * "
                nbutton.Text = dt.Rows(i).Item(1).ToString
                nbutton.CssClass = "selected_button"
            Else
                nbutton.Text = dt.Rows(i).Item(1).ToString
            End If
            nbutton.OnClientClick = "window.document.forms[0].AR_CODE.value=""" & gU.jsHTMLEncode(dt.Rows(i).Item(0).ToString) & """;window.document.forms[0].RK_CODE.value="""""
            'AddHandler nbutton.Click, AddressOf WHCODEClick

            tcAR_CODE.Controls.Add(nbutton)

            colCount += 1

            If colCount = 5 Then
                Dim lit As New Literal
                lit.Text = "<br/>"
                tcAR_CODE.Controls.Add(lit)
                colCount = 0
            End If
        Next
        trAR_CODE.Cells.Add(tcAR_CODE)

        REM Generate Rack
        colCount = 0
        Dim tclbltrRK_CODE As New TableCell()
        tclbltrRK_CODE.CssClass = "LabelTD"
        tclbltrRK_CODE.Style.Add("width", "20%")
        Dim RK_CODElabel As New Label
        RK_CODElabel.Text = "Rack"
        tclbltrRK_CODE.Controls.Add(RK_CODElabel)
        trRK_CODE.Cells.Add(tclbltrRK_CODE)

        Dim tcRK_CODE As New TableCell()
        tcRK_CODE.Style.Add("width", "80%")

        paP = New GlobalDBFunc.DBCmdPara
        srch_sql = "select RK_CODE, RK_NAME from WMS_WH_RACK WHERE WH_CODE = " & paP.AP(WH_CODE.Value) & " and FL_NUM = '" & FL_NUM.Value & "' AND AR_CODE = " & paP.AP(AR_CODE.Value) & " AND IMP_CODE = " & paP.AP(Session("IMP_CODE"))
        srch_sql += " order by RK_NAME"

        dt = gDB.getDataTable(srch_sql, , , , paP)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nbutton As New Button
            nbutton.Width = 75
            nbutton.CssClass = "loc_button"
            nbutton.ID = "RK_CODE_" & dt.Rows(i).Item(0).ToString
            If RK_CODE.Value = dt.Rows(i).Item(0).ToString Then
                'nbutton.Text = dt.Rows(i).Item(1).ToString & " * "
                nbutton.Text = dt.Rows(i).Item(1).ToString
                nbutton.CssClass = "selected_button"
            Else
                nbutton.Text = dt.Rows(i).Item(1).ToString
            End If
            nbutton.OnClientClick = "window.document.forms[0].RK_CODE.value=""" & gU.jsHTMLEncode(dt.Rows(i).Item(0).ToString) & """"
            'AddHandler nbutton.Click, AddressOf WHCODEClick

            tcRK_CODE.Controls.Add(nbutton)

            colCount += 1

            If colCount = 5 Then
                Dim lit As New Literal
                lit.Text = "<br/>"
                tcRK_CODE.Controls.Add(lit)
                colCount = 0
            End If
        Next
        trRK_CODE.Cells.Add(tcRK_CODE)

        REM Generate BIN
        Dim tclbltrBN_CODE As New TableCell()
        tclbltrBN_CODE.CssClass = "LabelTD"
        tclbltrBN_CODE.Style.Add("width", "20%")
        Dim BN_CODElabel As New Label
        BN_CODElabel.Text = "BIN"
        tclbltrBN_CODE.Controls.Add(BN_CODElabel)
        trBN_CODE.Cells.Add(tclbltrBN_CODE)

        Dim tcBN_CODE As New TableCell()
        tcBN_CODE.Style.Add("width", "80%")

        REM Create dt for BIN
        Dim nDT As New DataTable

        paP = New GlobalDBFunc.DBCmdPara
        srch_sql = "select IsNull(RK_X,'0')RK_X, IsNUll(RK_Y,'0')RK_Y from WMS_WH_RACK WHERE WH_CODE = " & paP.AP(WH_CODE.Value) & " and FL_NUM = '" & FL_NUM.Value & "' AND AR_CODE = " & paP.AP(AR_CODE.Value) & " AND RK_CODE = " & paP.AP(RK_CODE.Value) & " AND IMP_CODE = " & paP.AP(Session("IMP_CODE"))
        dt = gDB.getDataTable(srch_sql, , , , paP)
        If dt.Rows.Count > 0 Then
            For i As Integer = 1 To gU.decodeNull(dt.Rows(0).Item(0), 0)
                nDT.Columns.Add("COLUMN" & i)
            Next
            For i As Integer = 1 To gU.decodeNull(dt.Rows(0).Item(1), 0)
                nDT.Rows.Add()
            Next
        End If

        paP = New GlobalDBFunc.DBCmdPara
        srch_sql = "select BN_CODE, ISNULL(BN_X,0)BN_X, ISNULL(BN_Y,'0')BN_Y, BN_STATUS from WMS_WH_BIN WHERE WH_CODE = " & paP.AP(WH_CODE.Value) & " and FL_NUM = '" & FL_NUM.Value & "' AND AR_CODE = " & paP.AP(AR_CODE.Value) & " AND RK_CODE = " & paP.AP(RK_CODE.Value) & " AND IMP_CODE = " & paP.AP(Session("IMP_CODE"))
        srch_sql += " order by BN_CODE"

        dt = gDB.getDataTable(srch_sql, , , , paP)
        For i As Integer = 0 To dt.Rows.Count - 1
            If gU.decodeNull(dt.Rows(i).Item(2), 0) > 0 And gU.decodeNull(dt.Rows(i).Item(1), 0) > 0 Then
                If dt.Rows(i).Item("BN_STATUS").ToString = "INACTIVE" Then
                    nDT.Rows(dt.Rows(i).Item(2) - 1)("COLUMN" & dt.Rows(i).Item(1)) = ""
                Else
                    nDT.Rows(dt.Rows(i).Item(2) - 1)("COLUMN" & dt.Rows(i).Item(1)) = dt.Rows(i).Item(0).ToString
                End If
            End If
        Next
        Dim newTable As New Table
        For i As Integer = 0 To nDT.Rows.Count - 1
            Dim newRow As New TableRow
            For i1 As Integer = 0 To nDT.Columns.Count - 1
                Dim newCell As New TableCell
                Dim nbutton As New HtmlInputButton
                'nbutton.ID = "BN_CODE_" & nDT.Rows(i).Item(i1).ToString
                If nDT.Rows(i).Item(i1).ToString = "" Then
                    nbutton.Value = "        "
                    nbutton.Attributes.Add("class", "inactive_button")
                    nbutton.Disabled = True
                Else
                    If BN_CODE.Value = nDT.Rows(i).Item(i1).ToString Then
                        'nbutton.Value = nDT.Rows(i).Item(i1).ToString & " * "
                        nbutton.Value = nDT.Rows(i).Item(i1).ToString
                        nbutton.Attributes.Add("class", "selected_button")
                    Else
                        nbutton.Attributes.Add("class", "loc_button")
                        nbutton.Value = nDT.Rows(i).Item(i1).ToString
                    End If
                End If
                Dim locCode As String
                Dim codeValue As String
                'locCode = WH_CODE.Value & "" & Right("0" & FL_NUM.Value.ToString, 2) & "" & AR_CODE.Value & "" & RK_CODE.Value & "" & BN_CODE.Value
                Dim tempstr As String = ""
                'locCode = WH_CODE.Value & "" & FL_NUM.Value.ToString.PadLeft(2, "0") & "" & AR_CODE.Value.PadLeft(3, "0") & "" & RK_CODE.Value.ToString.PadLeft(4, "0") & "" & nDT.Rows(i).Item(i1).ToString.PadLeft(3, "0")
                locCode = WH_CODE.Value & "" & If(FL_NUM.Value.ToString.Length >= 2, FL_NUM.Value.ToString.Substring(0, 2), "00") & "" & If(AR_CODE.Value.ToString.Length >= 2, AR_CODE.Value.ToString.Substring(0, 2), "00") & "" & If(RK_CODE.Value.ToString.Length >= 2, RK_CODE.Value.ToString.Substring(0, 2), "00") & "" & If(nDT.Rows(i).Item(i1).ToString.Length >= 2, nDT.Rows(i).Item(i1).ToString.Substring(0, 2), "00")
                If nDT.Rows(i).Item(i1).ToString <> "" Then
                    Select Case FUN_CODE.Value
                        Case "OP_SRSRV"
                            Dim selectSQL As String = " SELECT (BN_LENGTH * BN_WIDTH) as bn_area, (BN_LENGTH * BN_WIDTH * BN_DEPTH) /1000000 as bn_cbm, BN_UTILIZATION_TYPE " &
                                                      " FROM WMS_WH_BIN WHERE (WH_CODE = '" & gU.dbEncode(WH_CODE.Value) & "') AND (AR_CODE = '" & gU.dbEncode(AR_CODE.Value) & "') AND (RK_CODE = '" & gU.dbEncode(RK_CODE.Value) & "') AND (FL_NUM = isnull('" & FL_NUM.Value & "','00')) AND (BN_CODE = '" & gU.dbEncode(nDT.Rows(i).Item(i1).ToString.Trim) & "') "

                            Dim tempDT As DataTable = gDB.getDataTable(selectSQL)

                            tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(0)) & "').innerHTML = '" & gU.jsHTMLEncode(locCode) & "';"
                            tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(1)) & "').value = '" & gU.jsHTMLEncode(locCode) & "';"

                            If tempDT.Rows.Count > 0 Then
                                tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(2)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(tempDT.Rows(0).Item("bn_cbm").ToString.Trim, "0")) & "';"
                                tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(3)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(tempDT.Rows(0).Item("bn_area").ToString.Trim, "0")) & "';"
                                tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(4)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(tempDT.Rows(0).Item("BN_UTILIZATION_TYPE").ToString.Trim, "CBM")) & "';"

                                Dim cbmInput As String = pItemArray(4)
                                Dim cbmPercent As Double = 0

                                If pItemArray(5) <> "" And gU.isDecimal(pItemArray(5)) Then
                                    Select Case tempDT.Rows(0).Item("BN_UTILIZATION_TYPE").ToString.Trim
                                        Case "AREA"
                                            If tempDT.Rows(0).Item("bn_area").ToString.Trim <> "" AndAlso gU.isDecimal(tempDT.Rows(0).Item("bn_area").ToString.Trim) Then
                                                If CDbl(tempDT.Rows(0).Item("bn_area").ToString.Trim) <> 0 Then
                                                    cbmPercent = Format((CDbl(pItemArray(5)) / CDbl(tempDT.Rows(0).Item("bn_area").ToString.Trim)) * 100, "0.000000")

                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(6)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(tempDT.Rows(0).Item("bn_area").ToString.Trim, "0")) & "';"
                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(7)) & "').innerHTML = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(tempDT.Rows(0).Item("bn_area").ToString.Trim, "0")) & "';"

                                                End If
                                            End If

                                        Case "CBM"
                                            If tempDT.Rows(0).Item("bn_cbm").ToString.Trim <> "" AndAlso gU.isDecimal(tempDT.Rows(0).Item("bn_cbm").ToString.Trim) Then
                                                If CDbl(tempDT.Rows(0).Item("bn_cbm").ToString.Trim) <> 0 Then
                                                    cbmPercent = Format((CDbl(pItemArray(5)) / CDbl(tempDT.Rows(0).Item("bn_cbm").ToString.Trim)) * 100, "0.000000")

                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(6)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(cbmPercent, "0")) & "';"
                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(7)) & "').innerHTML = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(cbmPercent, "0")) & "';"
                                                End If
                                            End If

                                        Case Else
                                            If tempDT.Rows(0).Item("bn_cbm").ToString.Trim <> "" AndAlso gU.isDecimal(tempDT.Rows(0).Item("bn_cbm").ToString.Trim) Then
                                                If CDbl(tempDT.Rows(0).Item("bn_cbm").ToString.Trim) <> 0 Then
                                                    cbmPercent = Format((CDbl(pItemArray(5)) / CDbl(tempDT.Rows(0).Item("bn_cbm").ToString.Trim)) * 100, "0.000000")

                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(6)) & "').value = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(cbmPercent, "0")) & "';"
                                                    tempstr &= "window.opener.document.getElementById('" & gU.jsHTMLEncode(pItemArray(7)) & "').innerHTML = '" & gU.jsHTMLEncode(gU.decodeNullOrEmpty(cbmPercent, "0")) & "';"
                                                End If
                                            End If

                                    End Select
                                End If


                            End If

                        Case Else
                            For i2 = 0 To pItemArray.Count - 1
                                Dim ptempItemArray As String()
                                ptempItemArray = Split(pItemArray(i2), "|")

                                If ptempItemArray.Count = 2 Then
                                    If ptempItemArray(1) = "L" Then
                                        tempstr = tempstr & "window.opener." & gU.jsHTMLEncode(ptempItemArray(0)) & ".innerHTML = '" & gU.jsHTMLEncode(locCode) & "';"
                                    ElseIf ptempItemArray(1) = "Q" Then
                                        tempstr = tempstr & "window.opener.document." & gU.jsHTMLEncode(pForm.Value) & "." & gU.jsHTMLEncode(ptempItemArray(0)) & ".value = " &
                                                                                    "'" & gU.jsHTMLEncode(getItemLocQty(sc.Value, itemCode.Value, packkey.Value, locCode)) & "';"
                                    ElseIf ptempItemArray(1) = "QL" Then
                                        tempstr = tempstr & "window.opener." & gU.jsHTMLEncode(ptempItemArray(0)) & ".innerHTML = " &
                                                                                    "'" & gU.jsHTMLEncode(getItemLocQty(sc.Value, itemCode.Value, packkey.Value, locCode)) & "';"
                                    Else
                                        tempstr = tempstr & "window.opener.document." & gU.jsHTMLEncode(pForm.Value) & "." & gU.jsHTMLEncode(ptempItemArray(0)) & ".value = '" & gU.jsHTMLEncode(locCode) & "';"
                                    End If
                                ElseIf ptempItemArray.Count = 3 Then
                                    If ptempItemArray(2) = "WH" Then
                                        codeValue = WH_CODE.Value
                                    ElseIf ptempItemArray(2) = "FL" Then
                                        codeValue = FL_NUM.Value
                                    ElseIf ptempItemArray(2) = "AR" Then
                                        codeValue = AR_CODE.Value
                                    ElseIf ptempItemArray(2) = "RK" Then
                                        codeValue = RK_CODE.Value
                                    ElseIf ptempItemArray(2) = "BN" Then
                                        codeValue = nDT.Rows(i).Item(i1).ToString()
                                    End If
                                    If ptempItemArray(1) = "L" Then
                                        tempstr = tempstr & "window.opener." & gU.jsHTMLEncode(ptempItemArray(0)) & ".innerHTML = '" & gU.jsHTMLEncode(codeValue) & "';"
                                    Else
                                        tempstr = tempstr & "window.opener.document." & gU.jsHTMLEncode(pForm.Value) & "." & gU.jsHTMLEncode(ptempItemArray(0)) & ".value = '" & gU.jsHTMLEncode(codeValue) & "';"
                                    End If
                                Else
                                    tempstr = tempstr & "window.opener.document." & gU.jsHTMLEncode(pForm.Value) & "." & gU.jsHTMLEncode(pItemArray(i2)) & ".value = '" & gU.jsHTMLEncode(locCode) & "';"
                                End If
                            Next

                    End Select


                    tempstr = tempstr & "window.open('','_self');window.close();"
                End If
                nbutton.Attributes.Add("onclick", "window.document.forms[0].BN_CODE.value=""" & gU.jsHTMLEncode(nDT.Rows(i).Item(i1).ToString) & """;" & tempstr & "")
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

        trBN_CODE.Cells.Add(tcBN_CODE)
    End Sub

    Private Function getItemLocQty(ByVal storercode As String, ByVal ic As String, ByVal pk As String, ByVal locCode As String) As String
        Dim nString As String = "select iloc_bal_qty, iloc_bal_cbm, iloc_bal_kg " & _
                                "from wms_item_loc_bal where " & _
                                "imp_code = '" & Session("IMP_CODE") & "' " & _
                                "and storer_code = '" & storercode & "' " & _
                                "and itm_code = '" & ic & "' " & _
                                "and pack_key = '" & pk & "' " & _
                                "and iloc_loc = '" & locCode & "'"

        Dim qtytbl As New DataTable

        qtytbl = gDB.getDataTable(nString)

        If qtytbl.Rows.Count > 0 Then
            Return gU.decodeEmptyCInt(qtytbl.Rows(0).Item("iloc_bal_qty").ToString, "0").ToString
        Else
            Return "0"
        End If
    End Function
End Class
