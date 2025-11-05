Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing


Partial Class MASTER_IM_wh_spec
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As New AccessRightUtils
    Private imp_code As String = ""

    Private DDFORMAT As String = ""


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

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            ViewState("storer_code") = ""
            ViewState("itm_code") = ""
            ViewState("pack_key") = ""
            ViewState("dt") = Nothing

            ViewState("storer_code") = Server.UrlDecode(Request("STORER_CODE"))
            ViewState("itm_code") = Server.UrlDecode(Request("ITM_CODE"))
            ViewState("pack_key") = Server.UrlDecode(Request("pack_key"))

            If ViewState("itm_code") = "" Then Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "", "window.close();", True)

            BindGV()
        Else
            If GridView1.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    DirectCast(GridView1.Rows(i).FindControl("dsp_IW_PREF_LOC1"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("IW_PREF_LOC1"), HiddenField).Value
                    DirectCast(GridView1.Rows(i).FindControl("dsp_IW_PREF_LOC2"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("IW_PREF_LOC2"), HiddenField).Value
                    DirectCast(GridView1.Rows(i).FindControl("dsp_IW_PICK_LOC"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("IW_PICK_LOC"), HiddenField).Value

                Next
            End If
        End If

        ar.hideForm(Me)

    End Sub

    Protected Sub BindGV()
        Dim selectSQL As String = ""
        Dim dt As DataTable
        Dim lstorer_code As String = ""
        Dim litm_code As String = ""
        Dim lpack_key As String = ""

        lstorer_code = ViewState("storer_code")
        litm_code = ViewState("itm_code")
        lpack_key = ViewState("pack_key")

        selectSQL = "SELECT 'U' as mFlag, IMP_CODE, ITM_CODE, WH_CODE, IW_REORD_QTY, IW_PREF_LOC1, IW_PREF_LOC2, IW_PICK_LOC, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB " & _
                    "FROM WMS_ITEM_WH " & _
                    "WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                    "AND ITM_CODE='" & gU.dbEncode(litm_code) & "' " & _
                    "AND pack_key='" & gU.dbEncode(lpack_key) & "' " & _
                    "and storer_code='" & lstorer_code & "' " & _
                    "order by WH_CODE "

        dt = gDB.getDataTable(selectSQL)


        GridView1.DataSource = dt
        ViewState("dt") = dt

        GridView1.DataBind()

    End Sub

    Private Function validateAll() As Boolean
        Dim tempDT As DataTable = ViewState("dt")
        Dim checkWH_list As String = ""
        Dim Cwh_code As String = ""
        If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                If tempDT.Rows(i).Item("mFlag") <> "D" Then
                    If DirectCast(GridView1.Rows(i).FindControl("WH_CODE"), DropDownList).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select Warehouse.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("WH_CODE"), DropDownList).Focus()
                        Else
                            uiFun.displayMsg(Me, "", "Please select Warehouse.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("WH_CODE"), DropDownList).Focus()
                        End If
                        Return False
                    Else
                        Cwh_code = DirectCast(GridView1.Rows(i).FindControl("WH_CODE"), DropDownList).SelectedValue

                        If gU.inList(checkWH_list, Cwh_code) Then

                            uiFun.displayMsg(Me, "", "Duplicate Warehouse Record.\nPlease Delete Duplicate record.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("WH_CODE"), DropDownList).Focus()
                            Return False

                        Else
                            checkWH_list = gU.appendToList(checkWH_list, Cwh_code)
                        End If
                    End If

                    If Not gU.isDecimal(DirectCast(GridView1.Rows(i).FindControl("IW_REORD_QTY"), TextBox).Text) Then
                        uiFun.displayMsg(Me, "", "Re-order Qty must be valid number!", Session("gLang"))
                        DirectCast(GridView1.Rows(i).FindControl("IW_REORD_QTY"), TextBox).Focus()
                        Return False
                    End If

                End If
            Next
        Else
            uiFun.displayMsg(Me, "", "No record has been added.", Session("gLang"))
            Return False
        End If

        Return True

    End Function

    Protected Sub Save(Optional ByVal flag As String = "")
        Dim detailDT As DataTable = ViewState("dt")
        Dim gConn As SqlConnection
        Dim selectSQL As String = ""
        Dim updateSQL As String = ""

        Dim successFlag As Boolean = False

        If validateAll() Then
            If cU.gfBuildDataTableforGridView(detailDT, GridView1, True) Then
                gConn = gDB.getConnection()
                Dim transaction As SqlTransaction = Nothing
                Dim paP As GlobalDBFunc.DBCmdPara
                Try
                    Dim lstorer_code As String = ""
                    Dim litm_code As String = ""
                    Dim lpack_key As String = ""

                    lstorer_code = ViewState("storer_code")
                    litm_code = ViewState("itm_code")
                    lpack_key = ViewState("pack_key")


                    gConn = gDB.getConnection()
                    transaction = gConn.BeginTransaction()

                    If Not detailDT Is Nothing AndAlso detailDT.Rows.Count > 0 Then
                        For i = 0 To detailDT.Rows.Count - 1
                            Select Case detailDT.Rows(i).Item("mFlag").ToString.Trim
                                Case "N"
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = "INSERT INTO WMS_ITEM_WH " & _
                                                "	(IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, WH_CODE, IW_REORD_QTY, IW_PREF_LOC1, IW_PREF_LOC2, IW_PICK_LOC, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB) " & _
                                                "VALUES (" & paP.AP(Session("IMP_CODE")) & ", " & paP.AP(lstorer_code) & "," & paP.AP(ViewState("itm_code")) & "," & paP.AP(lpack_key) & "," & paP.AP(detailDT.Rows(i).Item("wh_code").ToString.Trim) & "," & _
                                                paP.AP(detailDT.Rows(i).Item("IW_REORD_QTY").ToString.Trim, SqlDbType.Decimal) & "," & paP.AP(detailDT.Rows(i).Item("IW_PREF_LOC1").ToString.Trim) & "," & paP.AP(detailDT.Rows(i).Item("IW_PREF_LOC2").ToString.Trim) & "," & _
                                                paP.AP(detailDT.Rows(i).Item("IW_PICK_LOC").ToString.Trim) & "," & paP.AP(Session("usr_id")) & ",getdate(),getdate()," & paP.AP(Session("usr_id")) & ") "

                                Case "U"
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = "Update wms_item_WH set " & _
                                                "IW_REORD_QTY=" & paP.AP(detailDT.Rows(i).Item("IW_REORD_QTY").ToString.Trim, SqlDbType.Decimal) & "," & _
                                                "IW_PREF_LOC1=" & paP.AP(detailDT.Rows(i).Item("IW_PREF_LOC1").ToString.Trim) & ", IW_PREF_LOC2=" & paP.AP(detailDT.Rows(i).Item("IW_PREF_LOC2").ToString.Trim) & ", IW_PICK_LOC=" & paP.AP(detailDT.Rows(i).Item("IW_PICK_LOC").ToString.Trim) & ", SYS_LUB=" & paP.AP(Session("usr_id")) & ", SYS_LUD=getdate() " & _
                                                "Where IMP_CODE = " & paP.AP(Session("IMP_CODE")) & " " & _
                                                "and itm_code=" & paP.AP(ViewState("itm_code")) & " " & _
                                                "AND WH_CODE=" & paP.AP(detailDT.Rows(i).Item("WH_CODE").ToString.Trim) & " " & _
                                                "and pack_key=" & paP.AP(lpack_key) & " " & _
                                                "and storer_code=" & paP.AP(lstorer_code)
                                Case "D"
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = "Delete from wms_item_wh " & _
                                                "Where IMP_CODE = " & paP.AP(Session("IMP_CODE")) & " " & _
                                                "and itm_code=" & paP.AP(ViewState("itm_code")) & " " & _
                                                "AND WH_CODE=" & paP.AP(detailDT.Rows(i).Item("WH_CODE").ToString.Trim) & " " & _
                                                "and pack_key=" & paP.AP(lpack_key) & " " & _
                                                "and storer_code=" & paP.AP(lstorer_code)
                            End Select
                            'Dim tempsql As String = gDB.getCmdSql(updateSQL, paP)
                            If updateSQL <> "" Then gDB.amendData(updateSQL, gConn, transaction, paP)
                        Next
                    End If

                    transaction.Commit()
                    successFlag = True
                Catch ex As Exception
                    If Not transaction Is Nothing Then
                        transaction.Rollback()
                        transaction = Nothing
                    End If

                    Response.Write(ex.Message)
                    uiFun.displayMsg(Me, "1008", "", Session("gLang"))
                Finally
                    If gConn IsNot Nothing Then
                        If gConn.State = ConnectionState.Open Then
                            gConn.Close()
                            gConn.Dispose()
                        End If
                    End If
                End Try

                If successFlag Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                    Call BindGV()
                End If
            End If
        End If
    End Sub


    Protected Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        Dim tempDT As DataTable

        tempDT = ViewState("dt")

        If Not tempDT Is Nothing Then
            If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then
                Dim newRow As DataRow = tempDT.NewRow

                newRow.Item("mFlag") = "N"
                tempDT.Rows.Add(newRow)

                tempDT.AcceptChanges()

                ViewState("dt") = tempDT
                GridView1.DataSource = tempDT
                GridView1.DataBind()
            End If
        End If

    End Sub

    Protected Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click
        Call Save()
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("IW_REORD_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "IW_REORD_QTY").ToString.Trim
                uiFun.load_dropdown(CType(e.Row.FindControl("WH_CODE"), DropDownList), "SELECT WH_CODE, WH_NAME from WMS_WAREHOUSE ORDER by WH_NAME", "WH_CODE", "WH_NAME", , , DataBinder.Eval(e.Row.DataItem, "WH_CODE").ToString.Trim)

                CType(e.Row.FindControl("IW_PREF_LOC1"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "IW_PREF_LOC1").ToString.Trim
                CType(e.Row.FindControl("IW_PREF_LOC2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "IW_PREF_LOC2").ToString.Trim
                CType(e.Row.FindControl("IW_PICK_LOC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "IW_PICK_LOC").ToString.Trim


                CType(e.Row.FindControl("dsp_IW_PREF_LOC1"), Label).Text = DataBinder.Eval(e.Row.DataItem, "IW_PREF_LOC1").ToString.Trim
                CType(e.Row.FindControl("dsp_IW_PREF_LOC2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "IW_PREF_LOC2").ToString.Trim
                CType(e.Row.FindControl("dsp_IW_PICK_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "IW_PICK_LOC").ToString.Trim
                CType(e.Row.FindControl("Image_Loc_LookUp1"), Image).Attributes.Add("onclick", "LocLookUp(document.getElementById('" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "').value, '" & CType(e.Row.FindControl("dsp_IW_PREF_LOC1"), Label).ClientID & "', '" & CType(e.Row.FindControl("IW_PREF_LOC1"), HiddenField).ClientID & "', '" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "')")
                CType(e.Row.FindControl("Image_Loc_LookUp2"), Image).Attributes.Add("onclick", "LocLookUp(document.getElementById('" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "').value, '" & CType(e.Row.FindControl("dsp_IW_PREF_LOC2"), Label).ClientID & "', '" & CType(e.Row.FindControl("IW_PREF_LOC2"), HiddenField).ClientID & "', '" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "')")
                CType(e.Row.FindControl("Image_Loc_LookUp3"), Image).Attributes.Add("onclick", "LocLookUp(document.getElementById('" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "').value, '" & CType(e.Row.FindControl("dsp_IW_PICK_LOC"), Label).ClientID & "', '" & CType(e.Row.FindControl("IW_PICK_LOC"), HiddenField).ClientID & "', '" & CType(e.Row.FindControl("WH_CODE"), DropDownList).ClientID & "')")

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    CType(e.Row.FindControl("WH_CODE"), DropDownList).Enabled = False
                    CType(e.Row.FindControl("IW_REORD_QTY"), TextBox).Enabled = False
                    CType(e.Row.FindControl("Image_Loc_LookUp1"), Image).Visible = False
                    CType(e.Row.FindControl("Image_Loc_LookUp2"), Image).Visible = False
                    CType(e.Row.FindControl("Image_Loc_LookUp3"), Image).Visible = False
                ElseIf DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "U" Then
                    CType(e.Row.FindControl("WH_CODE"), DropDownList).Enabled = False
                End If
        End Select
    End Sub

    Protected Sub GridView1_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim tempDT As DataTable = ViewState("dt")

        If tempDT.Rows(e.RowIndex).Item("mFlag") = "N" Then
            tempDT.Rows(e.RowIndex).Delete()
        Else
            tempDT.Rows(e.RowIndex).Item("mFlag") = "D"
        End If


        tempDT.AcceptChanges()

        ViewState("dt") = tempDT
        GridView1.DataSource = tempDT
        GridView1.DataBind()
    End Sub
End Class
