Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_SCS_Main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As New AccessRightUtils

    
    Private dt As New DataTable


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        REM ****************************
        REM Modify Access Right Here
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
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            'uiFun.load_dropdownBy_ColCode(Status, "WMS_COL_CODE.COLC_STATUS", Session("gSelectLabel"))
            uiFun.load_dropdown(colc_tabcol, "select COLC_CODE, COLC_ENG_VALUE from WMS_COL_CODE WHERE COLC_TABCOL = 'SCS_CODE_TYPE' ORDER BY COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))

        Else
            'dt = ViewState("dt")

        End If

        If Session("pagemode") = "N" Then
            DelBtn1.Visible = False


        End If


        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "System Code Setup"
            lbl_code.Text = "Code"
            lbl_code_type.Text = "Code type"
            lbl_colc_eng_value.Text = "Code Name"
            lbl_display_seq.Text = "Display Seq."
            lbl_status.Text = "Status"

            saveBtn2.Text = "Save"
            saveBtn2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;"
            DelBtn1.Attributes("onclick") = "if(!confirm(""This System Code will be deleted. Are you sure to delete?"")) return false;"


        ElseIf Session("gLang") = "C" Then
            'lheader.Text = "Exam Master Review"
            'lbl_diet_code.Text = "Diet Code:"
            lheader.Text = "System Code Setup"
            lbl_code.Text = "Code"
            lbl_code_type.Text = "Code type"
            lbl_colc_eng_value.Text = "Code Name"
            lbl_display_seq.Text = "Display Seq."
            lbl_status.Text = "Status"

            saveBtn2.Text = "Save"
            saveBtn2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;;"
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        'Status.CssClass = "REQUIRED"
        colc_code.CssClass = "REQUIRED"
        colc_tabcol.CssClass = "REQUIRED"


        REM **********************

        If Session("pagemode") = "N" Then
            '.CssClass = "REQUIRED"
            '.CssClass = "REQUIRED"
        Else
            '.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing
            ViewState("colc_code") = ""
            ViewState("colc_tabcol") = ""

            ViewState("colc_code") = Server.UrlDecode(Request("colc_code"))
            ViewState("colc_tabcol") = Server.UrlDecode(Request("colc_tabcol"))

            Call BindGV()
        Else

            If Session("pagemode") <> "N" Then
                Call setControltoReadOnly(colc_code)
                Call setControltoReadOnly(colc_tabcol)
            End If
        End If

        Dim c_code, c_tabol As String

        If ViewState("colc_code") <> "" Then

            c_code = ViewState("colc_code")
            c_tabol = ViewState("colc_tabcol")

        Else

            c_code = Server.UrlDecode(Request("colc_code"))
            c_tabol = Server.UrlDecode(Request("colc_tabcol"))

        End If


        ar.hideForm(Me)


    End Sub


    Private Function validateAll(Optional ByVal cFlag As String = "") As Boolean

        If colc_display_seq.Text.Trim <> "" Then
            If Not IsNumeric(colc_display_seq.Text.Trim) Then
                uiFun.displayMsg(Me, "", lbl_display_seq.Text & " must be numeric!", Session("gLang"))
                Return False
            End If
        End If


        If colc_code.Text = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_code.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_code.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        Else
            If Not gU.isAlphaNumueric(colc_code.Text.ToString.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lbl_code.Text & " must be alphanumeric!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lbl_code.Text & "必須為英文字母及數字!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If colc_tabcol.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_code_type.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_code_type.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If Status.SelectedValue = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_status.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_status.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If


        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim sql_string As String = ""
        Dim up_sql As String = ""

        Dim gConn As SqlConnection
        Dim c_code As String = ""
        Dim c_tabcol As String = ""


        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable



        If ViewState("colc_code") <> "" Then

            c_code = ViewState("colc_code")
            c_tabcol = ViewState("colc_tabcol")

        Else
            c_code = Server.UrlDecode(Request("colc_code"))
            c_tabcol = Server.UrlDecode(Request("colc_tabcol"))

        End If

        If validateAll() Then


            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            'Start a local transaction

            Try
                If Session("pagemode") = "N" Then

                    dupSQL = "select 1 from WMS_COL_CODE " & _
                             "where colc_code=" & gU.convdbNVCData(gU.dbEncode(colc_code.Text.ToString.Trim)) & _
                             " AND colc_tabcol=" & gU.convdbNVCData(gU.dbEncode(colc_tabcol.SelectedValue))

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then

                        'Response.Write("ADD NEW")


                        sql_string = "Insert into WMS_COL_CODE (" & _
                                     "colc_tabcol, colc_code, colc_eng_value, colc_display_seq, colc_status)" & _
                                     " values ('" & gU.dbEncode(colc_tabcol.SelectedValue) & "' , '" & _
                                               gU.dbEncode(colc_code.Text) & "', '" & _
                                               gU.dbEncode(colc_eng_value.Text) & "', '" & _
                                               gU.dbEncode(colc_display_seq.Text) & "', '" & _
                                               gU.dbEncode(Status.SelectedValue) & "')"


                        If sql_string <> "" Then

                            gDB.amendData(sql_string, gConn, transaction)


                        End If

                        'transaction.Commit()
                        ViewState("colc_tabcol") = colc_tabcol.SelectedValue
                        ViewState("colc_code") = colc_code.Text

                    Else


                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        uiFun.displayMsg(Me, "", "Duplicate System Code is found !!", Session("gLang"))

                        Exit Sub

                    End If
                Else


                    sql_string = "update WMS_COL_CODE set " & _
                                     "colc_status='" & gU.dbEncode(Status.SelectedValue) & "', " & _
                                     "colc_eng_value='" & gU.dbEncode(colc_eng_value.Text.Trim) & "', " & _
                                     " colc_display_seq='" & gU.dbEncode(colc_display_seq.Text.Trim) & "' " & _
                                 "where colc_code='" & gU.dbEncode(c_code) & "'" & _
                                     " AND colc_tabcol='" & gU.dbEncode(c_tabcol) & "'"

                    If sql_string <> "" Then

                        gDB.amendData(sql_string, gConn, transaction)


                    End If


                End If



                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If

                transaction.Commit()

                Call BindGV()

                ' End If
            Catch ex As Exception

                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If
                Response.Write(sql_string & "<br><br>")
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
            'uiFun.displayMsg(Me, "", STATUS.SelectedValue, Session("gLang"))

        End If
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim cat_sql As String = ""
        Dim cat_dt As New DataTable
        Dim dt As New DataTable

        Dim c_code As String = ""
        Dim c_tabcol As String = ""

        REM **********************
        REM Modify Here
        REM Primary Key Session


        If ViewState("colc_tabcol") <> "" Then

            c_code = ViewState("colc_code")
            c_tabcol = ViewState("colc_tabcol")

        Else

            c_code = Server.UrlDecode(Request("colc_code"))
            c_tabcol = Server.UrlDecode(Request("colc_tabcol"))

        End If


        REM **********************

        'AD_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            'AD_STATUS.Text = "NEW"
            'AD_CODE.ForeColor = Drawing.Color.Red



            REM **********************
        Else
            DelBtn1.Visible = True


            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select * from WMS_COL_CODE where " & _
                        " colc_tabcol =" & gU.convdbNVCData(gU.dbEncode(c_tabcol)) & _
                        " AND colc_code =" & gU.convdbNVCData(gU.dbEncode(c_code))

            'Response.Write(SQLString)
            dt = gDB.getDataTable(SQLString)


            If dt.Rows.Count > 0 Then

                colc_code.Text = dt.Rows(0).Item("colc_code").ToString
                colc_tabcol.SelectedValue = dt.Rows(0).Item("colc_tabcol").ToString
                colc_display_seq.Text = dt.Rows(0).Item("colc_display_seq").ToString
                colc_eng_value.Text = dt.Rows(0).Item("colc_eng_value").ToString
                Status.SelectedValue = dt.Rows(0).Item("colc_status").ToString

                Call setControltoReadOnly(colc_code)
                Call setControltoReadOnly(colc_tabcol)

                REM **********************

            End If

            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        'Response.Write(SQLString)

        REM **********************
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub delbtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DelBtn1.Click
        Call delete_SCS()
    End Sub


    Protected Sub delete_SCS()
        Dim sql_string As String = ""
        Dim c_code As String = ""
        Dim c_tabcol As String = ""
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable
        Dim gConn As SqlConnection



        Dim del_success As String = "N"

        If ViewState("colc_tabcol") <> "" Then
            c_tabcol = ViewState("colc_tabcol")
            c_code = ViewState("colc_code")
        Else
            c_code = Server.UrlDecode(Request("colc_code"))
            c_tabcol = Server.UrlDecode(Request("colc_tabcol"))
        End If


        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        'Start a local transaction
        Try

            sql_string = "Delete from WMS_COL_CODE where colc_tabcol='" & gU.dbEncode(c_tabcol) & _
                         "' AND colc_code ='" & gU.dbEncode(c_code) & "'"


            If sql_string <> "" Then
                'Response.Write(sql_string & "<BR>")
                'Response.Write(itemsql)
                gDB.amendData(sql_string, gConn, transaction)


                del_success = "Y"
            End If
            transaction.Commit()



        Catch ex As Exception
            transaction.Rollback()
            Response.Write(sql_string & "<br><br>")
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

        If del_success = "Y" Then
            Response.Redirect("scs_main.aspx?mode=N")
        End If


    End Sub


    Private Sub setControltoReadOnly(ByRef ctl As Control)
        Select Case TypeName(ctl)

            Case "TextBox"
                Dim tempLabel As New Label

                tempLabel.ID = "DL_" & ctl.ID
                tempLabel.Style.Add("width", "auto")
                tempLabel.Font.Size = DirectCast(ctl, TextBox).Font.Size

                If DirectCast(ctl, TextBox).Text <> "" Then
                    tempLabel.Text = DirectCast(ctl, TextBox).Text
                Else
                    tempLabel.Text = ""
                End If

                ctl.Visible = False

                If Not TypeOf ctl.Parent Is HtmlForm AndAlso ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                    ctl.Parent.Controls.Add(tempLabel)
                Else
                    'DirectCast(ctl, TextBox).Visible = True
                    'DirectCast(ctl, TextBox).Enabled = False
                End If

            Case "DropDownList"

                Dim tempLabel As New Label
                tempLabel.ID = "DL_" & ctl.ID
                tempLabel.Style.Add("width", "auto")
                tempLabel.Font.Size = DirectCast(ctl, DropDownList).Font.Size

                If DirectCast(ctl, DropDownList).SelectedValue <> "" Then
                    tempLabel.Text = DirectCast(ctl, DropDownList).SelectedItem.Text
                Else
                    tempLabel.Text = ""
                End If

                ctl.Visible = False

                If Not TypeOf ctl.Parent Is HtmlForm AndAlso ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                    ctl.Parent.Controls.Add(tempLabel)
                Else
                    'ctl.Visible = True
                    'DirectCast(ctl, DropDownList).Enabled = False
                End If
        End Select
    End Sub

End Class