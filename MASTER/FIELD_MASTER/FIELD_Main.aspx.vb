Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_FIELD_Main
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
            ViewState("pagemode") = Nothing
            ViewState("pagemode") = Request("mode")

            ViewState("STORER_CODE") = ""
            ViewState("STORER_CODE") = Request("STORER_CODE")


        Else
            'dt = ViewState("dt")

        End If

        If ViewState("pagemode") = "N" Then



        End If


        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Customization Field setup"


            saveBtn2.Text = "Save"
            saveBtn2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;"



        ElseIf Session("gLang") = "C" Then
            'lheader.Text = "Exam Master Review"
            'lbl_diet_code.Text = "Diet Code:"
            lheader.Text = "Customization Field setup"


            saveBtn2.Text = "Save"
            saveBtn2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;;"
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        'Status.CssClass = "REQUIRED"



        REM **********************

        If ViewState("pagemode") = "N" Then
            '.CssClass = "REQUIRED"
            '.CssClass = "REQUIRED"
        Else
            '.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            If ViewState("STORER_CODE") <> "" Then
                STORER_NAME.Text = DB.getValueFromSQL("select STO_NAME from wms_storer where storer_code='" & gU.dbEncode(ViewState("STORER_CODE")) & "' AND IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "'")
            End If

            Call BindGV()
        Else

        End If

        ar.hideForm(Me)

    End Sub


    Private Function validateAll(Optional ByVal cFlag As String = "") As Boolean

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
        Dim gConn As SqlConnection
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim updateSQL As String = ""
        Dim lSTORER_CODE As String = ""
        Dim successFlag As Boolean = False

        lSTORER_CODE = ViewState("STORER_CODE")

        If validateAll() Then


            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            'Start a local transaction

            Try

                If ViewState("pagemode") = "N" Then

                Else
                    If Not field_list Is Nothing AndAlso field_list.Items.Count > 0 Then
                        For i = 0 To field_list.Items.Count - 1
                            updateSQL = ""
                            If DirectCast(field_list.Items(i).FindControl("FLDO_FIELD_OPTION"), CheckBox).Checked Then
                                paP = New GlobalDBFunc.DBCmdPara

                                Select Case FUN_CODE.SelectedValue
                                    Case "INQ_001", "RPT_SB"
                                        If ViewState("option_dt") IsNot Nothing Then
                                            Dim lDRow As DataRow() = DirectCast(ViewState("option_dt"), DataTable).Select("imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                                                                                                     "and storer_code = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                                                                                                     "and fun_code = '" & gU.dbEncode(FUN_CODE.SelectedValue) & "' " & _
                                                                                                                     "and fldo_field_name = '" & gU.dbEncode(DirectCast(field_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value) & "'")

                                            If lDRow.Count > 0 Then
                                                Dim row As DataRow = lDRow(0)

                                                updateSQL = ""
                                                updateSQL += "merge into wms_field_option opt "
                                                updateSQL += "using ("
                                                updateSQL += "select "
                                                updateSQL += paP.AP(row("imp_code").ToString) & " as imp_code, "
                                                updateSQL += paP.AP(row("storer_code").ToString) & " as storer_code, "
                                                updateSQL += paP.AP(row("fun_code").ToString) & " as fun_code, "
                                                updateSQL += paP.AP(row("fldo_field_name").ToString) & " as fldo_field_name, "
                                                updateSQL += paP.AP(row("fldo_field_desc").ToString) & " as fldo_field_desc, "
                                                updateSQL += paP.AP(row("fldo_table_name").ToString) & " as fldo_table_name, "
                                                updateSQL += paP.AP(row("fldo_field_ref").ToString) & " as fldo_field_ref, "
                                                updateSQL += paP.AP("Y") & " as fldo_field_option, "
                                                updateSQL += paP.AP(row("fldo_display_seq").ToString) & " as fldo_display_seq, "
                                                updateSQL += paP.AP(row("fldo_status").ToString) & " as fldo_status, "
                                                updateSQL += paP.AP(row("fldo_master_table_name").ToString) & " as fldo_master_table_name, "
                                                updateSQL += paP.AP(row("fldo_filed_ctrl").ToString) & " as fldo_filed_ctrl, "
                                                updateSQL += paP.AP(row("fldo_field_type").ToString) & " as fldo_field_type "
                                                updateSQL += ") src "
                                                updateSQL += "on ( "
                                                updateSQL += "opt.imp_code = src.imp_code "
                                                updateSQL += "and opt.storer_code = src.storer_code "
                                                updateSQL += "and opt.fun_code = src.fun_code "
                                                updateSQL += "and opt.fldo_field_name = src.fldo_field_name "
                                                updateSQL += ") "
                                                updateSQL += "when matched then "
                                                updateSQL += "update set opt.fldo_field_option = src.fldo_field_option "
                                                updateSQL += "when not matched then "
                                                updateSQL += "insert ("
                                                updateSQL += "imp_code, storer_code, fun_code, "
                                                updateSQL += "fldo_field_name, fldo_field_desc, "
                                                updateSQL += "fldo_table_name, fldo_field_ref, "
                                                updateSQL += "fldo_field_option, fldo_display_seq, "
                                                updateSQL += "fldo_status, fldo_master_table_name, "
                                                updateSQL += "fldo_filed_ctrl, fldo_field_type "
                                                updateSQL += ") values ("
                                                updateSQL += "src.imp_code, src.storer_code, src.fun_code, "
                                                updateSQL += "src.fldo_field_name, src.fldo_field_desc, "
                                                updateSQL += "src.fldo_table_name, src.fldo_field_ref, "
                                                updateSQL += "src.fldo_field_option, src.fldo_display_seq, "
                                                updateSQL += "src.fldo_status, src.fldo_master_table_name, "
                                                updateSQL += "src.fldo_filed_ctrl, src.fldo_field_type "
                                                updateSQL += ")"
                                            End If
                                        End If

                                    Case Else
                                        updateSQL = "Update wms_field_option set FLDO_FIELD_OPTION='Y' " & _
                                            "where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(lSTORER_CODE) & " AND FUN_CODE=" & paP.AP(FUN_CODE.SelectedValue) & _
                                            " AND FLDO_FIELD_NAME=" & paP.AP(DirectCast(field_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value)

                                End Select
                            Else
                                paP = New GlobalDBFunc.DBCmdPara

                                Select Case FUN_CODE.SelectedValue
                                    Case "INQ_001", "RPT_SB"
                                        If ViewState("option_dt") IsNot Nothing Then
                                            Dim lDRow As DataRow() = DirectCast(ViewState("option_dt"), DataTable).Select("imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                                                                                                                                                "and storer_code = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                                                                                                                                                "and fun_code = '" & gU.dbEncode(FUN_CODE.SelectedValue) & "' " & _
                                                                                                                                                                "and fldo_field_name = '" & gU.dbEncode(DirectCast(field_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value) & "'")

                                            If lDRow.Count > 0 Then
                                                Dim row As DataRow = lDRow(0)

                                                updateSQL = ""
                                                updateSQL += "merge into wms_field_option opt "
                                                updateSQL += "using ("
                                                updateSQL += "select "
                                                updateSQL += paP.AP(row("imp_code").ToString) & " as imp_code, "
                                                updateSQL += paP.AP(row("storer_code").ToString) & " as storer_code, "
                                                updateSQL += paP.AP(row("fun_code").ToString) & " as fun_code, "
                                                updateSQL += paP.AP(row("fldo_field_name").ToString) & " as fldo_field_name, "
                                                updateSQL += paP.AP(row("fldo_field_desc").ToString) & " as fldo_field_desc, "
                                                updateSQL += paP.AP(row("fldo_table_name").ToString) & " as fldo_table_name, "
                                                updateSQL += paP.AP(row("fldo_field_ref").ToString) & " as fldo_field_ref, "
                                                updateSQL += paP.AP("N") & " as fldo_field_option, "
                                                updateSQL += paP.AP(row("fldo_display_seq").ToString) & " as fldo_display_seq, "
                                                updateSQL += paP.AP(row("fldo_status").ToString) & " as fldo_status, "
                                                updateSQL += paP.AP(row("fldo_master_table_name").ToString) & " as fldo_master_table_name, "
                                                updateSQL += paP.AP(row("fldo_filed_ctrl").ToString) & " as fldo_filed_ctrl, "
                                                updateSQL += paP.AP(row("fldo_field_type").ToString) & " as fldo_field_type "
                                                updateSQL += ") src "
                                                updateSQL += "on ( "
                                                updateSQL += "opt.imp_code = src.imp_code "
                                                updateSQL += "and opt.storer_code = src.storer_code "
                                                updateSQL += "and opt.fun_code = src.fun_code "
                                                updateSQL += "and opt.fldo_field_name = src.fldo_field_name "
                                                updateSQL += ") "
                                                updateSQL += "when matched then "
                                                updateSQL += "update set opt.fldo_field_option = src.fldo_field_option "
                                                updateSQL += "when not matched then "
                                                updateSQL += "insert ("
                                                updateSQL += "imp_code, storer_code, fun_code, "
                                                updateSQL += "fldo_field_name, fldo_field_desc, "
                                                updateSQL += "fldo_table_name, fldo_field_ref, "
                                                updateSQL += "fldo_field_option, fldo_display_seq, "
                                                updateSQL += "fldo_status, fldo_master_table_name, "
                                                updateSQL += "fldo_filed_ctrl, fldo_field_type "
                                                updateSQL += ") values ("
                                                updateSQL += "src.imp_code, src.storer_code, src.fun_code, "
                                                updateSQL += "src.fldo_field_name, src.fldo_field_desc, "
                                                updateSQL += "src.fldo_table_name, src.fldo_field_ref, "
                                                updateSQL += "src.fldo_field_option, src.fldo_display_seq, "
                                                updateSQL += "src.fldo_status, src.fldo_master_table_name, "
                                                updateSQL += "src.fldo_filed_ctrl, src.fldo_field_type "
                                                updateSQL += ")"
                                            End If
                                        End If

                                    Case Else
                                        updateSQL = "Update wms_field_option set FLDO_FIELD_OPTION='N' " & _
                                            "where imp_code=" & paP.AP(Session("imp_code")) & " AND  STORER_CODE=" & paP.AP(lSTORER_CODE) & " AND FUN_CODE=" & paP.AP(FUN_CODE.SelectedValue) & _
                                            " AND FLDO_FIELD_NAME=" & paP.AP(DirectCast(field_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value)
                                End Select
                            End If

                            If updateSQL <> "" Then
                                gDB.amendData(updateSQL, gConn, transaction, paP)
                            End If
                        Next
                    End If




                    If Not report_list Is Nothing AndAlso report_list.Items.Count > 0 Then
                        For i = 0 To report_list.Items.Count - 1
                            updateSQL = ""
                            If DirectCast(report_list.Items(i).FindControl("FLDO_FIELD_OPTION"), CheckBox).Checked Then
                                paP = New GlobalDBFunc.DBCmdPara

                                Select Case FUN_CODE.SelectedValue
                                    Case "INQ_001", "RPT_SB"
                                        If ViewState("option_dt3") IsNot Nothing Then
                                            Dim lDRow As DataRow() = DirectCast(ViewState("option_dt3"), DataTable).Select("imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                                                                                                     "and storer_code = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                                                                                                     "and fun_code = '" & gU.dbEncode(FUN_CODE.SelectedValue) & "' " & _
                                                                                                                     "and fldo_field_name = '" & gU.dbEncode(DirectCast(report_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value) & "'")

                                            If lDRow.Count > 0 Then
                                                Dim row As DataRow = lDRow(0)

                                                updateSQL = ""
                                                updateSQL += "merge into wms_field_option opt "
                                                updateSQL += "using ("
                                                updateSQL += "select "
                                                updateSQL += paP.AP(row("imp_code").ToString) & " as imp_code, "
                                                updateSQL += paP.AP(row("storer_code").ToString) & " as storer_code, "
                                                updateSQL += paP.AP(row("fun_code").ToString) & " as fun_code, "
                                                updateSQL += paP.AP(row("fldo_field_name").ToString) & " as fldo_field_name, "
                                                updateSQL += paP.AP(row("fldo_field_desc").ToString) & " as fldo_field_desc, "
                                                updateSQL += paP.AP(row("fldo_table_name").ToString) & " as fldo_table_name, "
                                                updateSQL += paP.AP(row("fldo_field_ref").ToString) & " as fldo_field_ref, "
                                                updateSQL += paP.AP("Y") & " as fldo_field_option, "
                                                updateSQL += paP.AP(row("fldo_display_seq").ToString) & " as fldo_display_seq, "
                                                updateSQL += paP.AP(row("fldo_status").ToString) & " as fldo_status, "
                                                updateSQL += paP.AP(row("fldo_master_table_name").ToString) & " as fldo_master_table_name, "
                                                updateSQL += paP.AP(row("fldo_filed_ctrl").ToString) & " as fldo_filed_ctrl, "
                                                updateSQL += paP.AP(row("fldo_field_type").ToString) & " as fldo_field_type "
                                                updateSQL += ") src "
                                                updateSQL += "on ( "
                                                updateSQL += "opt.imp_code = src.imp_code "
                                                updateSQL += "and opt.storer_code = src.storer_code "
                                                updateSQL += "and opt.fun_code = src.fun_code "
                                                updateSQL += "and opt.fldo_field_name = src.fldo_field_name "
                                                updateSQL += ") "
                                                updateSQL += "when matched then "
                                                updateSQL += "update set opt.fldo_field_option = src.fldo_field_option "
                                                updateSQL += "when not matched then "
                                                updateSQL += "insert ("
                                                updateSQL += "imp_code, storer_code, fun_code, "
                                                updateSQL += "fldo_field_name, fldo_field_desc, "
                                                updateSQL += "fldo_table_name, fldo_field_ref, "
                                                updateSQL += "fldo_field_option, fldo_display_seq, "
                                                updateSQL += "fldo_status, fldo_master_table_name, "
                                                updateSQL += "fldo_filed_ctrl, fldo_field_type "
                                                updateSQL += ") values ("
                                                updateSQL += "src.imp_code, src.storer_code, src.fun_code, "
                                                updateSQL += "src.fldo_field_name, src.fldo_field_desc, "
                                                updateSQL += "src.fldo_table_name, src.fldo_field_ref, "
                                                updateSQL += "src.fldo_field_option, src.fldo_display_seq, "
                                                updateSQL += "src.fldo_status, src.fldo_master_table_name, "
                                                updateSQL += "src.fldo_filed_ctrl, src.fldo_field_type "
                                                updateSQL += ")"
                                            End If
                                        End If
                                End Select
                            Else
                                paP = New GlobalDBFunc.DBCmdPara

                                Select Case FUN_CODE.SelectedValue
                                    Case "INQ_001", "RPT_SB"
                                        If ViewState("option_dt3") IsNot Nothing Then
                                            Dim lDRow As DataRow() = DirectCast(ViewState("option_dt3"), DataTable).Select("imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                                                                                                                                                "and storer_code = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                                                                                                                                                "and fun_code = '" & gU.dbEncode(FUN_CODE.SelectedValue) & "' " & _
                                                                                                                                                                "and fldo_field_name = '" & gU.dbEncode(DirectCast(report_list.Items(i).FindControl("FLDO_FIELD_NAME"), HiddenField).Value) & "'")

                                            If lDRow.Count > 0 Then
                                                Dim row As DataRow = lDRow(0)

                                                updateSQL = ""
                                                updateSQL += "merge into wms_field_option opt "
                                                updateSQL += "using ("
                                                updateSQL += "select "
                                                updateSQL += paP.AP(row("imp_code").ToString) & " as imp_code, "
                                                updateSQL += paP.AP(row("storer_code").ToString) & " as storer_code, "
                                                updateSQL += paP.AP(row("fun_code").ToString) & " as fun_code, "
                                                updateSQL += paP.AP(row("fldo_field_name").ToString) & " as fldo_field_name, "
                                                updateSQL += paP.AP(row("fldo_field_desc").ToString) & " as fldo_field_desc, "
                                                updateSQL += paP.AP(row("fldo_table_name").ToString) & " as fldo_table_name, "
                                                updateSQL += paP.AP(row("fldo_field_ref").ToString) & " as fldo_field_ref, "
                                                updateSQL += paP.AP("N") & " as fldo_field_option, "
                                                updateSQL += paP.AP(row("fldo_display_seq").ToString) & " as fldo_display_seq, "
                                                updateSQL += paP.AP(row("fldo_status").ToString) & " as fldo_status, "
                                                updateSQL += paP.AP(row("fldo_master_table_name").ToString) & " as fldo_master_table_name, "
                                                updateSQL += paP.AP(row("fldo_filed_ctrl").ToString) & " as fldo_filed_ctrl, "
                                                updateSQL += paP.AP(row("fldo_field_type").ToString) & " as fldo_field_type "
                                                updateSQL += ") src "
                                                updateSQL += "on ( "
                                                updateSQL += "opt.imp_code = src.imp_code "
                                                updateSQL += "and opt.storer_code = src.storer_code "
                                                updateSQL += "and opt.fun_code = src.fun_code "
                                                updateSQL += "and opt.fldo_field_name = src.fldo_field_name "
                                                updateSQL += ") "
                                                updateSQL += "when matched then "
                                                updateSQL += "update set opt.fldo_field_option = src.fldo_field_option "
                                                updateSQL += "when not matched then "
                                                updateSQL += "insert ("
                                                updateSQL += "imp_code, storer_code, fun_code, "
                                                updateSQL += "fldo_field_name, fldo_field_desc, "
                                                updateSQL += "fldo_table_name, fldo_field_ref, "
                                                updateSQL += "fldo_field_option, fldo_display_seq, "
                                                updateSQL += "fldo_status, fldo_master_table_name, "
                                                updateSQL += "fldo_filed_ctrl, fldo_field_type "
                                                updateSQL += ") values ("
                                                updateSQL += "src.imp_code, src.storer_code, src.fun_code, "
                                                updateSQL += "src.fldo_field_name, src.fldo_field_desc, "
                                                updateSQL += "src.fldo_table_name, src.fldo_field_ref, "
                                                updateSQL += "src.fldo_field_option, src.fldo_display_seq, "
                                                updateSQL += "src.fldo_status, src.fldo_master_table_name, "
                                                updateSQL += "src.fldo_filed_ctrl, src.fldo_field_type "
                                                updateSQL += ")"
                                            End If
                                        End If
                                End Select
                            End If

                            If updateSQL <> "" Then
                                gDB.amendData(updateSQL, gConn, transaction, paP)
                            End If
                        Next
                    End If
                End If


                If ViewState("pagemode") = "N" Then
                    ViewState("pagemode") = ""

                    REM **********************
                    REM Modify Here

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If

                transaction.Commit()

                successFlag = True

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

            If successFlag Then
                Call BindGV()
            End If

        End If
    End Sub

    Protected Sub BindGV()
        Dim selectSQL As String = ""
        Dim dt As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim paP As GlobalDBFunc.DBCmdPara

        Dim lSTORER_CODE As String = ""

        REM **********************
        REM Modify Here
        REM Primary Key Session

        lSTORER_CODE = ViewState("STORER_CODE")


        If FUN_CODE.SelectedValue <> "" Then
            paP = New GlobalDBFunc.DBCmdPara

            Select Case FUN_CODE.SelectedValue
                Case "INQ_001", "RPT_SB"
                    selectSQL = ""
                    selectSQL += "select ISNULL(wms_field_option.imp_code, " & paP.AP(Session("imp_code")) & ") as imp_code, "
                    selectSQL += "ISNULL(wms_field_option.storer_code, " & paP.AP(lSTORER_CODE) & ") as storer_code, "
                    selectSQL += "wms_search_col.fun_code, "
                    selectSQL += "wms_search_col.cold_tabcol as fldo_field_name,"
                    selectSQL += "wms_search_col.srcl_label as fldo_field_desc, "
                    selectSQL += "ISNULL(wms_field_option.fldo_field_option, 'Y') as fldo_field_option,  "
                    selectSQL += "wms_search_col.srcl_list_seq as fldo_display_seq, "
                    selectSQL += "wms_field_option.fldo_table_name, "
                    selectSQL += "wms_field_option.fldo_field_ref,  "
                    selectSQL += "ISNULL(wms_field_option.fldo_status, 'A') as fldo_status,  "
                    selectSQL += "wms_field_option.fldo_master_table_name, wms_field_option.fldo_filed_ctrl, wms_field_option.fldo_field_type "
                    selectSQL += "from wms_search_col left outer join wms_field_option on "
                    selectSQL += "wms_search_col.cold_tabcol = wms_field_option.fldo_field_name "
                    selectSQL += "and wms_search_col.fun_code = wms_field_option.fun_code "
                    selectSQL += "and wms_field_option.imp_code= " & paP.AP(Session("imp_code")) & " "
                    selectSQL += "and wms_field_option.storer_code= " & paP.AP(lSTORER_CODE) & " "
                    selectSQL += "where wms_search_col.fun_code= " & paP.AP(FUN_CODE.SelectedValue) & " "
                    selectSQL += "and wms_search_col.srcl_list_seq <> '0' "
                    selectSQL += "order by wms_search_col.srcl_list_seq "

                Case Else
                    selectSQL = " SELECT wms_field_option.IMP_CODE,  wms_field_option.STORER_CODE,  wms_field_option.FUN_CODE, " & _
                        "  wms_field_option.FLDO_FIELD_NAME,  wms_field_option.FLDO_FIELD_DESC,  wms_field_option.FLDO_TABLE_NAME, " & _
                        "  wms_field_option.FLDO_FIELD_REF,  wms_field_option.FLDO_FIELD_OPTION,  wms_field_option.FLDO_DISPLAY_SEQ, " & _
                        "  wms_field_option.FLDO_STATUS,  wms_field_option.FLDO_MASTER_TABLE_NAME,  wms_field_option.FLDO_FILED_CTRL " & _
                        " FROM wms_field_option " & _
                        " Where wms_field_option.imp_code=" & paP.AP(Session("imp_code")) & " AND wms_field_option.Storer_code=" & paP.AP(lSTORER_CODE) & " AND wms_field_option.FUN_CODE=" & paP.AP(FUN_CODE.SelectedValue) & _
                        " AND ISNULL(FLDO_FIELD_TYPE, '') = '' " & _
                        " Order by wms_field_option.FLDO_TABLE_NAME, wms_field_option.FLDO_DISPLAY_SEQ, wms_field_option.FLDO_FIELD_DESC "
            End Select

            dt = gDB.getDataTable(selectSQL, , , , paP)


            If dt.Rows.Count > 0 Then
                field_list.DataSource = dt
                ViewState("option_dt") = dt
            Else
                field_list.DataSource = Nothing
                ViewState("option_dt") = Nothing
            End If

            field_list.DataBind()

            paP = New GlobalDBFunc.DBCmdPara

            Select Case FUN_CODE.SelectedValue
                Case "INQ_001", "RPT_SB"
                    exclude_list.DataSource = Nothing
                    ViewState("option_dt2") = Nothing
                    manTr.Visible = False

                Case Else
                    selectSQL = " SELECT wms_field_option.IMP_CODE,  wms_field_option.STORER_CODE,  wms_field_option.FUN_CODE, " & _
                        "  wms_field_option.FLDO_FIELD_NAME,  wms_field_option.FLDO_FIELD_DESC,  wms_field_option.FLDO_TABLE_NAME, " & _
                        "  wms_field_option.FLDO_FIELD_REF,  wms_field_option.FLDO_FIELD_OPTION,  wms_field_option.FLDO_DISPLAY_SEQ, " & _
                        "  wms_field_option.FLDO_STATUS,  wms_field_option.FLDO_MASTER_TABLE_NAME,  wms_field_option.FLDO_FILED_CTRL " & _
                        " FROM wms_field_option " & _
                        " Where wms_field_option.imp_code=" & paP.AP(Session("imp_code")) & " AND wms_field_option.Storer_code=" & paP.AP(lSTORER_CODE) & " AND wms_field_option.FUN_CODE=" & paP.AP(FUN_CODE.SelectedValue) & _
                        " AND ISNULL(FLDO_FIELD_TYPE, '') <> '' " & _
                        " Order by wms_field_option.FLDO_TABLE_NAME, wms_field_option.FLDO_DISPLAY_SEQ, wms_field_option.FLDO_FIELD_DESC "

                    dt2 = gDB.getDataTable(selectSQL, , , , paP)


                    If dt2.Rows.Count > 0 Then
                        exclude_list.DataSource = dt2
                        ViewState("option_dt2") = dt2
                    Else
                        exclude_list.DataSource = Nothing
                        ViewState("option_dt2") = Nothing
                    End If

                    manTr.Visible = True
            End Select

            exclude_list.DataBind()


            paP = New GlobalDBFunc.DBCmdPara

            Select Case FUN_CODE.SelectedValue
                Case "INQ_001", "RPT_SB"
                    selectSQL = ""
                    selectSQL += "select ISNULL(wms_field_option.imp_code, " & paP.AP(Session("imp_code")) & ") as imp_code, "
                    selectSQL += "ISNULL(wms_field_option.storer_code, " & paP.AP(lSTORER_CODE) & ") as storer_code, "
                    selectSQL += "wms_search_col.fun_code, "
                    selectSQL += "wms_search_col.cold_tabcol as fldo_field_name,"
                    selectSQL += "wms_search_col.srcl_label as fldo_field_desc, "
                    selectSQL += "ISNULL(wms_field_option.fldo_field_option, 'Y') as fldo_field_option,  "
                    selectSQL += "wms_search_col.srcl_list_seq as fldo_display_seq, "
                    selectSQL += "wms_field_option.fldo_table_name, "
                    selectSQL += "wms_field_option.fldo_field_ref,  "
                    selectSQL += "ISNULL(wms_field_option.fldo_status, 'A') as fldo_status,  "
                    selectSQL += "wms_field_option.fldo_master_table_name, wms_field_option.fldo_filed_ctrl, wms_field_option.fldo_field_type "
                    selectSQL += "from wms_search_col left outer join wms_field_option on "
                    selectSQL += "wms_search_col.cold_tabcol = wms_field_option.fldo_field_name "
                    selectSQL += "and wms_search_col.fun_code = wms_field_option.fun_code "
                    selectSQL += "and wms_field_option.imp_code= " & paP.AP(Session("imp_code")) & " "
                    selectSQL += "and wms_field_option.storer_code= " & paP.AP(lSTORER_CODE) & " "
                    selectSQL += "where wms_search_col.fun_code= " & paP.AP(FUN_CODE.SelectedValue) & " "
                    selectSQL += "and wms_search_col.srcl_list_seq = '0' "
                    selectSQL += "and wms_search_col.srcl_rpt_seq <> '0' "
                    selectSQL += "order by wms_search_col.srcl_rpt_seq "

                    dt3 = gDB.getDataTable(selectSQL, , , , paP)

                    If dt3.Rows.Count > 0 Then
                        report_list.DataSource = dt3
                        ViewState("option_dt3") = dt3
                    Else
                        report_list.DataSource = Nothing
                        ViewState("option_dt3") = Nothing
                    End If

                    rptTr.Visible = True
                Case Else
                    rptTr.Visible = False
            End Select

            report_list.DataBind()
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

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        Call BindGV()
    End Sub

    Protected Sub field_list_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataListItemEventArgs) Handles field_list.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                DirectCast(e.Item.FindControl("FLDO_FIELD_DESC"), Label).Text = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_DESC").ToString.Trim
                DirectCast(e.Item.FindControl("FLDO_FIELD_NAME"), HiddenField).Value = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_NAME").ToString.Trim

                If DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_OPTION").ToString.Trim = "Y" Then
                    DirectCast(e.Item.FindControl("FLDO_FIELD_OPTION"), CheckBox).Checked = True
                End If
        End Select
    End Sub

    Protected Sub exclude_list_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataListItemEventArgs) Handles exclude_list.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                DirectCast(e.Item.FindControl("FLDO_FIELD_DESC"), Label).Text = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_DESC").ToString.Trim
                DirectCast(e.Item.FindControl("FLDO_FIELD_NAME"), HiddenField).Value = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_NAME").ToString.Trim
        End Select
    End Sub

    Protected Sub report_list_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataListItemEventArgs) Handles report_list.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                DirectCast(e.Item.FindControl("FLDO_FIELD_DESC"), Label).Text = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_DESC").ToString.Trim
                DirectCast(e.Item.FindControl("FLDO_FIELD_NAME"), HiddenField).Value = DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_NAME").ToString.Trim

                If DataBinder.Eval(e.Item.DataItem, "FLDO_FIELD_OPTION").ToString.Trim = "Y" Then
                    DirectCast(e.Item.FindControl("FLDO_FIELD_OPTION"), CheckBox).Checked = True
                End If
        End Select
    End Sub

    Protected Sub FUN_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles FUN_CODE.SelectedIndexChanged
        Call BindGV()
    End Sub
End Class