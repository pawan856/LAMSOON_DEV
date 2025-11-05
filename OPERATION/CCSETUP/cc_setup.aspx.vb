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
        ar = New AccessRightUtils("OP_CCSETUP", Session("usr_id"), Me)

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        'If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
        '    Exit Sub
        'End If

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            ViewState("dt") = Nothing
            BindGV()
        End If

        ar.hideForm(Me)

    End Sub

    Protected Sub BindGV()
        Dim selectSQL As String = ""
        Dim dt As DataTable
        

        selectSQL = " SELECT 'U' as mflag, CCS_YEAR, CCS_PERIOD, convert(varchar,CCS_END_DATE,103) as CCS_END_DATE, sys_cb, sys_lub, sys_cd, sys_lud " & _
                    " FROM WMS_CC_SETUP "

        dt = gDB.getDataTable(selectSQL)

        GridView1.DataSource = dt
        ViewState("dt") = dt

        GridView1.DataBind()

    End Sub

    Private Function validateAll() As Boolean
        Dim tempDT As DataTable = ViewState("dt")
        Dim checkYR_list As String = ""
        Dim YrKey As String = ""
        If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                If tempDT.Rows(i).Item("mFlag") <> "D" Then

                    If DirectCast(GridView1.Rows(i).FindControl("CCS_YEAR"), DropDownList).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select Year.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("CCS_YEAR"), DropDownList).Focus()
                        Else
                            uiFun.displayMsg(Me, "", "Please select Year.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("CCS_YEAR"), DropDownList).Focus()
                        End If
                        Return False
                    End If

                    If DirectCast(GridView1.Rows(i).FindControl("CCS_PERIOD"), DropDownList).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select Period.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("CCS_PERIOD"), DropDownList).Focus()
                        Else
                            uiFun.displayMsg(Me, "", "Please select Period.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("CCS_PERIOD"), DropDownList).Focus()
                        End If
                        Return False
                    Else
                        YrKey = DirectCast(GridView1.Rows(i).FindControl("CCS_YEAR"), DropDownList).SelectedValue & DirectCast(GridView1.Rows(i).FindControl("CCS_PERIOD"), DropDownList).SelectedValue

                        If gU.inList(checkYR_list, YrKey) Then

                            uiFun.displayMsg(Me, "", "Duplicate Year/Period Combination.\nPlease Delete Duplicate record.", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("CCS_YEAR"), DropDownList).Focus()
                            Return False

                        Else
                            checkYR_list = gU.appendToList(checkYR_list, YrKey)
                        End If
                    End If

                    If DirectCast(GridView1.Rows(i).FindControl("CCS_END_DATE"), TextBox).Text = "" Then
                        uiFun.displayMsg(Me, "", "Please enter the end date!", Session("gLang"))
                        DirectCast(GridView1.Rows(i).FindControl("CCS_END_DATE"), TextBox).Focus()
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
                    gConn = gDB.getConnection()
                    transaction = gConn.BeginTransaction()

                    If Not detailDT Is Nothing AndAlso detailDT.Rows.Count > 0 Then
                        For i = 0 To detailDT.Rows.Count - 1
                            Select Case detailDT.Rows(i).Item("mFlag").ToString.Trim
                                Case "N"
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = " INSERT INTO WMS_CC_SETUP (CCS_YEAR, CCS_PERIOD, CCS_END_DATE, sys_cb, sys_lub, sys_cd, sys_lud) " & _
                                                " VALUES        (" & paP.AP(detailDT.Rows(i).Item("CCS_YEAR").ToString.Trim) & "," & paP.AP(detailDT.Rows(i).Item("CCS_PERIOD").ToString.Trim) & "," & gU.convdbDate(gU.dbEncode(detailDT.Rows(i).Item("CCS_END_DATE").ToString.Trim)) & "," & _
                                                paP.AP(Session("usr_id")) & "," & paP.AP(Session("usr_id")) & ",getdate(),getdate()) "
                                Case "U"
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = "update WMS_CC_SETUP set " & _
                                                "CCS_END_DATE=" & gU.convdbDate(gU.dbEncode(detailDT.Rows(i).Item("CCS_END_DATE").ToString.Trim)) & "," & _
                                                "sys_lud=getdate(), sys_lub=" & paP.AP(Session("usr_id")) & _
                                                " where ccs_year=" & paP.AP(detailDT.Rows(i).Item("CCS_YEAR").ToString.Trim) & " AND CCS_PERIOD=" & paP.AP(detailDT.Rows(i).Item("CCS_PERIOD").ToString.Trim)

                                Case "D"
                                    'paP = New GlobalDBFunc.DBCmdPara
                                    'updateSQL = "Delete from wms_item_wh " & _
                                    '            "Where IMP_CODE = " & paP.AP(Session("IMP_CODE")) & " " & _
                                    '            "and itm_code=" & paP.AP(ViewState("itm_code")) & " " & _
                                    '            "AND WH_CODE=" & paP.AP(detailDT.Rows(i).Item("WH_CODE").ToString.Trim) & " " & _
                                    '            "and pack_key=" & paP.AP(lpack_key) & " " & _
                                    '            "and storer_code=" & paP.AP(lstorer_code)
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

                Dim lItem As ListItem

                For i = 2013 To 2030
                    lItem = New ListItem
                    lItem.Value = i
                    lItem.Text = i

                    CType(e.Row.FindControl("CCS_YEAR"), DropDownList).Items.Add(lItem)
                Next

                CType(e.Row.FindControl("mFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    CType(e.Row.FindControl("CCS_YEAR"), DropDownList).SelectedValue = Now.Year
                Else
                    CType(e.Row.FindControl("CCS_YEAR"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "CCS_YEAR").ToString.Trim
                    CType(e.Row.FindControl("CCS_PERIOD"), DropDownList).Enabled = False
                    CType(e.Row.FindControl("CCS_YEAR"), DropDownList).Enabled = False

                End If

                CType(e.Row.FindControl("CCS_PERIOD"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "CCS_PERIOD").ToString.Trim
                CType(e.Row.FindControl("CCS_END_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CCS_END_DATE").ToString.Trim



        End Select
    End Sub
   
    Protected Sub btnSave2_Click(sender As Object, e As System.EventArgs) Handles btnSave2.Click
        Call Save()
    End Sub
End Class
