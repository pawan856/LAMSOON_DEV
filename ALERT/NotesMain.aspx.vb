Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Globalization

Partial Class ALERT_NotesMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then
            Select Case Session("gLang")
                Case "C"
                    lHeader.Text = "工單"
                    lbl_DOC_TYPE.Text = "貨單類型"
                    lbl_DOC_NO.Text = "單號"
                    lHeaderNew.Text = "新增工單"
                    btnPnlSave.Text = "儲存"
                    btnPnlClose.Text = "取消"
                    GridView1.EmptyDataText = "找不到工單"
                    lblHDGV.Text = "工單列表"
                    btnNew.Text = "新增工單"

                    lHeaderNew.Text = "工單內容"
                    lbl_TASK_PRIORITY.Text = "優先度"
                    lbl_TASK_SUBJECT.Text = "主題"
                    lbl_TASK_DETAILS.Text = "內容"
                    lbl_TASK_FROM.Text = "寄件者"
                    lbl_SYS_CD.Text = "寄件日期"

                Case Else
                    lHeader.Text = "Note"
                    lbl_DOC_TYPE.Text = "Order Type"
                    lbl_DOC_NO.Text = "Order No."
                    lHeaderNew.Text = "New Notes"
                    btnPnlSave.Text = "Save"
                    btnPnlClose.Text = "Cancel"
            End Select

            ViewState("DOC_TYPE") = ""
            ViewState("DOC_NO") = ""
            ViewState("STORE_CODE") = ""
            Session("NoteDT") = Nothing

            ViewState("DOC_TYPE") = Server.UrlDecode(Request("DOC_TYPE"))
            ViewState("DOC_NO") = Server.UrlDecode(Request("DOC_NO"))
            ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

            BindGV()
            changeLabel()

            uiFun.load_checkboxList(ALERT_TO, "Select USR_FNAME + ' ' + USR_SNAME + ' - ' + case when usr_type='S' then 'Staff' when usr_type='T' then 'Customer' else '' end as USER_NAME, usr_id FROM WMS_USER WHERE USR_STATUS = 'A'", "usr_id", "user_name")

        End If

        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            If ViewState("STORE_CODE") = "" Then
                Response.End()
            End If

            If Session("usr_pref_storer") <> ViewState("STORE_CODE") Then
                Response.End()
            End If
        End If

    End Sub

    Protected Sub BindGV()
        Dim dt As DataTable
        Dim selectSQL As String = ""
        Dim lstorer_code, ldoc_type, ldoc_no As String

        lstorer_code = ViewState("STORER_CODE")
        ldoc_type = ViewState("DOC_TYPE")
        ldoc_no = ViewState("DOC_NO")

        Dim paP As GlobalDBFunc.DBCmdPara

        Dim tempDT As DataTable

        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = "Select FUN_ENG_NAME, FUN_CHI_NAME FROM WMS_FUNCTION WHERE FUN_CODE=" & paP.AP(ldoc_type)

        tempDT = gDB.getDataTable(selectSQL, , , , paP)

        If tempDT.Rows.Count > 0 Then
            Select Case Session("gLang")
                Case "C"
                    DOC_TYPE.Text = tempDT.Rows(0).Item("FUN_CHI_NAME").ToString.Trim
                Case Else
                    DOC_TYPE.Text = tempDT.Rows(0).Item("FUN_ENG_NAME").ToString.Trim
            End Select

        End If
        DOC_NO.Text = ldoc_no


        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = " SELECT SYS_PK, SEQ, IMP_CODE, STORER_CODE, TASK_DEFAULT_YN, SYS_PROJ_NO, PROJ_NO, DOC_TYPE, DOC_NO, TASK_PRIORITY, STATUS, TASK_SUBJECT, TASK_DETAILS, TASK_FROM, TASK_ASSIGN_TO, " & _
                     " TASK_TARGET_DATE, TASK_DONE_DATE, WMS_NOTES.SYS_CB, WMS_NOTES.SYS_CD, WMS_NOTES.SYS_UB, WMS_NOTES.SYS_UD, TASK_TYPE, FUN_CODE, TASK_EMAIL_ALERT_YN, DOC_SYS_PK, TASK_CUST_VIEW_YN, CONV_OLD_ID_NO, CONV_BATCH_ID, " & _
                     " F_USER.USR_FNAME + ' ' +  F_USER.USR_SNAME as USR_NAME_FR " & _
                     " FROM WMS_NOTES " & _
                     " LEFT OUTER JOIN WMS_USER F_USER ON F_USER.USR_ID = TASK_FROM " & _
                     " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE =" & paP.AP(lstorer_code) & " AND DOC_TYPE=" & paP.AP(ldoc_type) & " AND DOC_NO=" & paP.AP(ldoc_no) & _
                     " order by SYS_CD DESC"

        dt = gDB.getDataTable(selectSQL, , , , paP)

        Session("NoteDT") = dt
        GridView1.DataSource = dt
        GridView1.DataBind()

    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here

        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "Posted By", "發起人")
        Call cU.newChangeGVLabel(GridView1, "Subject", "主題")
        Call cU.newChangeGVLabel(GridView1, "Details", "內容")
        Call cU.newChangeGVLabel(GridView1, "Date", "日期")
       
        REM **********************
    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case e.CommandName
            Case "VIEW"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                Dim sysPK As String = DirectCast(gvRow.FindControl("sys_pk"), HiddenField).Value
                BindPnl("V", sysPK)
            Case "REPLY"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                Dim sysPK As String = DirectCast(gvRow.FindControl("sys_pk"), HiddenField).Value
                BindPnl("R", sysPK)
        End Select
    End Sub

    Protected Sub GridView1_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
       
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("TASK_FROM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "USR_NAME_FR").ToString.Trim
                CType(e.Row.FindControl("TASK_SUBJECT"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TASK_SUBJECT").ToString.Trim
                CType(e.Row.FindControl("TASK_DETAILS"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TASK_DETAILS").ToString.Trim
                CType(e.Row.FindControl("SYS_CD"), Label).Text = DataBinder.Eval(e.Row.DataItem, "SYS_CD").ToString.Trim
                CType(e.Row.FindControl("SYS_PK"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "SYS_PK").ToString.Trim
        End Select
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(btnNew)



    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        BindPnl("N")
    End Sub


    Protected Sub ResetNote()
        TASK_PRIORITY_N.SelectedIndex = 0
        TASK_SUBJECT_N.Text = ""
        TASK_DETAILS.Text = ""
        vw_TASK_DETAILS.Text = ""
        vw_TASK_SUBJECT.Text = ""
        TASK_SUBJECT_N.Visible = True
        TASK_DETAILS.Visible = True

        If ALERT_TO.Items.Count > 0 Then
            For i = 0 To ALERT_TO.Items.Count - 1
                ALERT_TO.Items(i).Selected = False
            Next
        End If

    End Sub

    Protected Sub BindPnl(ByVal action As String, Optional ByVal skey As String = "")
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim tempDT As DataTable

        Select Case action
            Case "N"
                ResetNote()
                Alert_TO_TR.Visible = True
                CAction.Value = "N"
                TR_SYS_CD.Visible = False
                TASK_FROM.Text = Session("usr_id")
                TASK_PRIORITY_N.Attributes.Clear()
                pnlDtl_ModalPopupExtender.Show()
            Case "R"
                ResetNote()
                Alert_TO_TR.Visible = True
                CAction.Value = "N"

                paP = New GlobalDBFunc.DBCmdPara
                selectSQL = "Select WMS_NOTES.SYS_CD, PROJ_NO, TASK_FROM, TASK_SUBJECT, TASK_DETAILS, TASK_PRIORITY, F_USER.USR_FNAME + ' ' +  F_USER.USR_SNAME as USR_NAME_FR FROM WMS_NOTES LEFT OUTER JOIN WMS_USER F_USER ON F_USER.USR_ID = TASK_FROM WHERE SYS_PK=" & paP.AP(skey)
                tempDT = gDB.getDataTable(selectSQL, , , , paP)

                If tempDT.Rows.Count > 0 Then
                    TASK_SUBJECT_N.Text = "RE: " & tempDT.Rows(0).Item("TASK_SUBJECT").ToString.Trim
                    TASK_PRIORITY_N.SelectedValue = tempDT.Rows(0).Item("TASK_PRIORITY").ToString.Trim
                    TASK_PRIORITY_N.Attributes.Add("disabled", "true")
                    ALERT_TO.SelectedValue = tempDT.Rows(0).Item("TASK_FROM").ToString.Trim
                End If

                TR_SYS_CD.Visible = False
                TASK_FROM.Text = Session("usr_id")
                TASK_PRIORITY_N.Attributes.Clear()
                pnlDtl_ModalPopupExtender.Show()
            Case "V"
                'TASK_PRIORITY_N.Visible = False
                TASK_SUBJECT_N.Visible = False
                TASK_DETAILS.Visible = False

                TR_SYS_CD.Visible = True
                Alert_TO_TR.Visible = False

                CAction.Value = "V"

                If skey <> "" Then
                    paP = New GlobalDBFunc.DBCmdPara
                    selectSQL = "Select WMS_NOTES.SYS_CD, PROJ_NO, TASK_FROM, TASK_SUBJECT, TASK_DETAILS, TASK_PRIORITY, F_USER.USR_FNAME + ' ' +  F_USER.USR_SNAME as USR_NAME_FR FROM WMS_NOTES LEFT OUTER JOIN WMS_USER F_USER ON F_USER.USR_ID = TASK_FROM WHERE SYS_PK=" & paP.AP(skey)
                    tempDT = gDB.getDataTable(selectSQL, , , , paP)

                    If tempDT.Rows.Count > 0 Then
                        vw_TASK_DETAILS.Text = tempDT.Rows(0).Item("TASK_DETAILS").ToString.Trim
                        vw_TASK_SUBJECT.Text = tempDT.Rows(0).Item("TASK_SUBJECT").ToString.Trim
                        TASK_PRIORITY_N.SelectedValue = tempDT.Rows(0).Item("TASK_PRIORITY").ToString.Trim
                        TASK_PRIORITY_N.Attributes.Add("disabled", "true")
                        SYS_CD.Text = tempDT.Rows(0).Item("SYS_CD").ToString.Trim
                        TASK_FROM.Text = tempDT.Rows(0).Item("USR_NAME_FR").ToString.Trim

                        pnlDtl_ModalPopupExtender.Show()
                    Else
                        uiFun.displayMsgNew(MainUDP, "", "No Notes has been found!", Session("gLang"))
                    End If
                End If

        End Select
    End Sub

    Protected Sub btnPnlSave_Click(sender As Object, e As System.EventArgs) Handles btnPnlSave.Click
        SavePNL(CAction.Value)
    End Sub

    Protected Sub SavePNL(ByVal action As String)
        Dim updateSQL As String = ""
        Dim selectSQL As String = ""

        Dim lproj_no As String = ""

        Dim paP As GlobalDBFunc.DBCmdPara

        Dim lstorer_code, ldoc_type, ldoc_no As String

        lstorer_code = ViewState("STORER_CODE")
        ldoc_type = ViewState("DOC_TYPE")
        ldoc_no = ViewState("DOC_NO")

        Dim successFlag As Boolean = False

        Dim gConn As SqlConnection

        If validateAll() Then
            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction
            ' Start a local transaction
            Try
                Select Case action
                    Case "N"

                        Select Case ldoc_type
                            Case "IB_CO"
                                paP = New GlobalDBFunc.DBCmdPara
                                selectSQL = "SELECT CO_EDI_SIR_NO from WMS_CUST_ORDER WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(lstorer_code) & " AND CO_CODE=" & paP.AP(ldoc_no)
                                lproj_no = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSQL, gConn, transaction, paP), "")

                            Case "IB_DO"
                                paP = New GlobalDBFunc.DBCmdPara
                                selectSQL = "SELECT DO_EDI_SIR_NO from WMS_DELV_ORDER WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(lstorer_code) & " AND DO_CODE=" & paP.AP(ldoc_no)
                                lproj_no = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSQL, gConn, transaction, paP), "")


                            Case Else
                                lproj_no = ""
                        End Select

                        Dim sysPK As String = ""

                        sysPK = DB.getDocNo("NOTE", gConn, transaction)

                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = " INSERT INTO WMS_NOTES " & _
                                    " (SYS_PK, SEQ, IMP_CODE, STORER_CODE, TASK_DEFAULT_YN, SYS_PROJ_NO, PROJ_NO, DOC_TYPE, DOC_NO, TASK_PRIORITY, STATUS, TASK_SUBJECT, TASK_DETAILS, TASK_FROM, TASK_ASSIGN_TO,  " & _
                                    " TASK_TYPE, FUN_CODE, TASK_EMAIL_ALERT_YN, DOC_SYS_PK, TASK_CUST_VIEW_YN, " & _
                                    " SYS_CB, SYS_CD, SYS_UB, SYS_UD ) " & _
                                    " VALUES(" & _
                                    paP.AP(sysPK) & ",1," & paP.AP(Session("IMP_CODE")) & "," & paP.AP(lstorer_code) & ",'N',NULL," & paP.AP(lproj_no) & "," & paP.AP(ldoc_type) & "," & paP.AP(ldoc_no) & "," &
                                    paP.AP(TASK_PRIORITY_N.SelectedValue) & ",'N'," & paP.AP(TASK_SUBJECT_N.Text.Trim, SqlDbType.NVarChar) & "," & paP.AP(TASK_DETAILS.Text.Trim, SqlDbType.NVarChar) & ", " & paP.AP(Session("usr_id")) & ",NULL,'NOTE'," & _
                                    paP.AP(ldoc_type) & ",'N','1','N'," & _
                                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        gDB.amendData(updateSQL, gConn, transaction, paP)

                        If ALERT_TO.Items.Count > 0 Then
                            Dim asysPK As String = ""

                            For i = 0 To ALERT_TO.Items.Count - 1
                                If ALERT_TO.Items(i).Selected Then
                                    asysPK = DB.getDocNo("ALERT", gConn, transaction)

                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = " INSERT INTO WMS_ALERT " & _
                                                " (SYS_PK, ALRT_TYPE, ALRT_SEND_BY, FUN_CODE, IMP_CODE, STORER_CODE, SYS_DOC_NO, ALRT_REF_NO, ALRT_SUBJECT, ALRT_MSG, ALRT_TO_USERID, " & _
                                                " ALRT_SEND_DATE, DOC_SYS_PK, ALRT_PRIORITY, SYS_CB, SYS_CD, SYS_UB, SYS_UD)  " & _
                                                " VALUES (" & _
                                                paP.AP(asysPK) & ", 'LOGIN', " & paP.AP(Session("usr_id")) & "," & paP.AP(ldoc_type) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(lstorer_code) & "," & paP.AP(ldoc_no) & "," & _
                                                paP.AP(lproj_no) & "," & paP.AP(TASK_SUBJECT_N.Text.Trim) & "," & paP.AP(TASK_DETAILS.Text.Trim) & "," & paP.AP(ALERT_TO.Items(i).Value) & ",getdate(),'1'," & _
                                                paP.AP(TASK_PRIORITY_N.SelectedValue) & "," & _
                                                "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    gDB.amendData(updateSQL, gConn, transaction, paP)
                                End If
                            Next
                        End If
                End Select

                transaction.Commit()
                successFlag = True
            Catch ex As Exception
                transaction.Rollback()
                uiFun.displayMsgNew(MainUDP, "", ex.Message, Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        Else
            pnlDtl_ModalPopupExtender.Show()
        End If
        
        If successFlag Then
            Select Case Session("gLang")
                Case "C"
                    uiFun.displayMsgNew(MainUDP, "", "新增工單完成!", Session("gLang"))
                Case Else
                    uiFun.displayMsgNew(MainUDP, "", "New Note has been added", Session("gLang"))
            End Select

            BindGV()
        End If
    End Sub

    Private Function ValidateAll() As Boolean

        If String.IsNullOrWhiteSpace(TASK_DETAILS.Text) Then
            Select Case Session("gLang")
                Case "C"
                    uiFun.displayMsgNew(MainUDP, "", "請輸入工單內容!", Session("gLang"))
                Case Else
                    uiFun.displayMsgNew(MainUDP, "", "Please enter Note Details!", Session("gLang"))
            End Select

            TASK_DETAILS.Focus()
            Return False
        End If

        Return True
    End Function
End Class
