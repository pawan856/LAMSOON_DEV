Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Globalization

Partial Class ALERT_AlertMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Dim selectSQL As String
    Dim paP As GlobalDBFunc.DBCmdPara

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Session("usr_id") = "" Then
            Response.End()
        End If

        If Not IsPostBack Then
            Session("AlertDT") = Nothing

            Select Case Session("gLang")
                Case "C"
                    lHeader.Text = "訊息提示"
                    lbl_ALERT_TO.Text = "收件人"
                    lbl_KEY_WORD.Text = "關鍵字"
                    lbl_SENT_DATE.Text = "寄件日期"
                    btnRead.Text = "己閱及移除"
                    GridView1.EmptyDataText = "找不到訊息"
                    btnClose.Text = "關閉"
                Case Else
                    lHeader.Text = "Alert List"

                    lbl_ALERT_TO.Text = "Sent To"
                    lbl_KEY_WORD.Text = "Keywords"
                    lbl_SENT_DATE.Text = "Sent Date"
                    btnClose.Text = "Close"
            End Select

            Dim tempSQL As String

            tempSQL = " AND USR_ID='" & Session("usr_id") & "'"
            uiFun.load_dropdown(ALERT_TO, "SELECT USR_ID, USR_FNAME + ' ' +  USR_SNAME as USR_NAME FROM WMS_USER WHERE USR_STATUS='A'" & tempSQL, "USR_ID", "USR_NAME", , , Session("usr_id"), True)


            BindGV()
            
        End If
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)
        sm.RegisterAsyncPostBackControl(btnSearch)
        sm.RegisterAsyncPostBackControl(btnRead)
        sm.RegisterAsyncPostBackControl(SORT_BY)

    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here

        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "Posted By", "發起人")
        Call cU.newChangeGVLabel(GridView1, "Subject", "主題")
        Call cU.newChangeGVLabel(GridView1, "Details", "內容")

        REM **********************
    End Sub


    Protected Sub BindGV()
        Dim dt As DataTable
        Dim tempSQL As String = ""


        paP = New GlobalDBFunc.DBCmdPara
        If Not String.IsNullOrWhiteSpace(KEY_WORD.Text) Then
            tempSQL &= " AND (ALRT_SUBJECT LIKE '%" & gU.dbEncode(KEY_WORD.Text.Trim) & "%' or ALRT_MSG like '%" & gU.dbEncode(KEY_WORD.Text.Trim) & "%') "
        End If

        If Not String.IsNullOrWhiteSpace(DATE_FROM.Text) Then
            tempSQL &= " AND ALRT_SEND_DATE >= Convert(datetime," & paP.AP(DATE_FROM.Text) & ",103) "
        End If

        If Not String.IsNullOrWhiteSpace(DATE_TO.Text) Then
            tempSQL &= " AND ALRT_SEND_DATE < Convert(datetime," & paP.AP(DATE_TO.Text) & ",103) + 1"
        End If


        selectSQL = " SELECT SYS_PK, ALRT_TYPE, ALRT_SEND_BY, WMS_ALERT.FUN_CODE, FUN_NAME, WMS_ALERT.IMP_CODE, WMS_ALERT.STORER_CODE, SYS_DOC_NO, ALRT_REF_NO, ALRT_REPORT_REQ_KEY, ALRT_RPT_FORMAT, ALRT_SUBJECT, ALRT_MSG, ALRT_TO_USERID, " & _
                    " ALRT_EMAIL_TO, ALRT_EMAIL_CC, ALRT_SEND_DATE, WMS_ALERT.SYS_CB, WMS_ALERT.SYS_CD, WMS_ALERT.SYS_UB, WMS_ALERT.SYS_UD, CONV_OLD_ID_NO, CONV_BATCH_ID, DOC_SYS_PK, ALRT_TMP_KEY, ALRT_URL, ALRT_URL_DESC, ALRT_IMAGE, " & _
                    " ALRT_SEND_TO, ALRT_PRIORITY, FR_USER.USR_FNAME + ' ' +  FR_USER.USR_SNAME as ALRT_SEND_BY_NAME, TO_USER.USR_FNAME + ' ' +  TO_USER.USR_SNAME as ALRT_SEND_TO_NAME,  " & _
                    " Case when ALRT_PRIORITY='H' then 0 when ALRT_PRIORITY='M' then 1 when ALRT_PRIORITY='L' then 3 else 4 end as PRIORITY_SEQ, STO_NAME, FUN_ENG_NAME, FUN_CHI_NAME " & _
                    " FROM WMS_ALERT " & _
                    " LEFT OUTER JOIN WMS_USER FR_USER ON ALRT_SEND_BY = FR_USER.USR_ID " & _
                    " LEFT OUTER JOIN WMS_USER TO_USER ON ALRT_TO_USERID = TO_USER.USR_ID " & _
                    " LEFT OUTER JOIN WMS_STORER ON WMS_STORER.STORER_CODE = WMS_ALERT.STORER_CODE AND WMS_STORER.IMP_CODE = WMS_ALERT.IMP_CODE " & _
                    " LEFT OUTER JOIN WMS_FUNCTION ON WMS_ALERT.FUN_CODE = WMS_FUNCTION.FUN_CODE " & _
                    " WHERE ALRT_TO_USERID=" & paP.AP(ALERT_TO.SelectedValue) & tempSQL & _
                    " ORDER BY ALRT_SEND_DATE desc"

        dt = gDB.getDataTable(selectSQL, , , , paP)

        GridView1.DataSource = dt
        GridView1.DataBind()

        Session("AlertDT") = dt
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
          Select e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("sys_pk"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "sys_pk").ToString.Trim

                Select Case DataBinder.Eval(e.Row.DataItem, "ALRT_PRIORITY").ToString.Trim
                    Case "M"
                        CType(e.Row.FindControl("ALRT_PRIORITY"), Label).Text = "Medium"
                    Case "L"
                        CType(e.Row.FindControl("ALRT_PRIORITY"), Label).Text = "Low"
                    Case "H"
                        CType(e.Row.FindControl("ALRT_PRIORITY"), Label).Text = "High"
                    Case Else
                        CType(e.Row.FindControl("ALRT_PRIORITY"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_PRIORITY").ToString.Trim
                End Select

                CType(e.Row.FindControl("ALRT_SUBJECT"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_SUBJECT").ToString.Trim
                CType(e.Row.FindControl("ALRT_MSG"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_MSG").ToString.Trim
                CType(e.Row.FindControl("ALRT_TO_USERID"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_SEND_TO_NAME").ToString.Trim
                CType(e.Row.FindControl("ALRT_SEND_BY"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_SEND_BY_NAME").ToString.Trim
                CType(e.Row.FindControl("STO_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "STO_NAME").ToString.Trim

                Select Case Session("gLang")
                    Case "C"
                        CType(e.Row.FindControl("SYS_DOC_NO"), HyperLink).Text = DataBinder.Eval(e.Row.DataItem, "FUN_CHI_NAME").ToString.Trim & vbCrLf & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                    Case Else
                        CType(e.Row.FindControl("SYS_DOC_NO"), HyperLink).Text = DataBinder.Eval(e.Row.DataItem, "FUN_ENG_NAME").ToString.Trim & vbCrLf & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                End Select

                Dim redirectAdd As String = ""
                Select Case DataBinder.Eval(e.Row.DataItem, "FUN_CODE").ToString.Trim
                    Case "IB_RO"
                        redirectAdd = "../INBOUND/RO/ROMain.aspx?STORER_CODE=" & DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim & "&RO_CODE=" & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                    Case "IB_GR"
                        redirectAdd = "../INBOUND/GR/GRMain.aspx?STORER_CODE=" & DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim & "&GR_CODE=" & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                    Case "OB_CO"
                        redirectAdd = "../OUTBOUND/CO/COMain.aspx?STORER_CODE=" & DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim & "&CO_CODE=" & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                    Case "OB_DO"
                        redirectAdd = "../OUTBOUND/DO/DOMain.aspx?STORER_CODE=" & DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim & "&DO_CODE=" & DataBinder.Eval(e.Row.DataItem, "SYS_DOC_NO").ToString.Trim
                End Select

                CType(e.Row.FindControl("SYS_DOC_NO"), HyperLink).NavigateUrl = "#"
                CType(e.Row.FindControl("SYS_DOC_NO"), HyperLink).Attributes.Add("onclick", "javascript:goLocation('" & redirectAdd & "')")

                CType(e.Row.FindControl("ALRT_REF_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_REF_NO").ToString.Trim
                CType(e.Row.FindControl("ALRT_SEND_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ALRT_SEND_DATE").ToString.Trim

        End Select
    End Sub

    Protected Sub btnRead_Click(sender As Object, e As System.EventArgs) Handles btnRead.Click
        ReadAlert()
    End Sub

    Protected Sub SORT_BY_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles SORT_BY.SelectedIndexChanged
        SortDT(SORT_BY.SelectedValue)
    End Sub

    Protected Sub ReadAlert()
        Dim updateSQL As String = ""
        Dim sysPK As String = ""

        Dim successFlag As Boolean = False

        If GridView1.Rows.Count > 0 Then
            Dim gConn As SqlConnection

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction
            ' Start a local transaction
            Try
                Dim updateCount As Integer = 0

                For i = 0 To GridView1.Rows.Count - 1
                    If DirectCast(GridView1.Rows(i).FindControl("CheckYN"), CheckBox).Checked Then
                        sysPK = DirectCast(GridView1.Rows(i).FindControl("sys_pk"), HiddenField).Value

                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = " INSERT INTO WMS_ALERT_HISTORY " & _
                                    " (SYS_PK, ALRT_TYPE, ALRT_SEND_BY, FUN_CODE, FUN_NAME, IMP_CODE, STORER_CODE, SYS_DOC_NO, ALRT_REF_NO, ALRT_REPORT_REQ_KEY, ALRT_RPT_FORMAT, ALRT_SUBJECT, ALRT_MSG, ALRT_TO_USERID, " & _
                                    " ALRT_EMAIL_TO, ALRT_EMAIL_CC, ALRT_SEND_DATE, ALRT_READ_DATE, SYS_CB, SYS_CD, SYS_UB, SYS_UD, CONV_OLD_ID_NO, CONV_BATCH_ID, DOC_SYS_PK, ALRT_TMP_KEY, ALRT_URL, ALRT_URL_DESC, ALRT_IMAGE, ALRT_SEND_TO) " & _
                                    " (SELECT " & _
                                    " SYS_PK, ALRT_TYPE, ALRT_SEND_BY, FUN_CODE, FUN_NAME, IMP_CODE, STORER_CODE, SYS_DOC_NO, ALRT_REF_NO, ALRT_REPORT_REQ_KEY, ALRT_RPT_FORMAT, ALRT_SUBJECT, ALRT_MSG, ALRT_TO_USERID, " & _
                                    " ALRT_EMAIL_TO, ALRT_EMAIL_CC, ALRT_SEND_DATE,getdate(), SYS_CB, SYS_CD, SYS_UB, SYS_UD, CONV_OLD_ID_NO, CONV_BATCH_ID, DOC_SYS_PK, ALRT_TMP_KEY, ALRT_URL, ALRT_URL_DESC, ALRT_IMAGE, ALRT_SEND_TO " & _
                                    " FROM WMS_ALERT WHERE SYS_PK=" & paP.AP(sysPK) & ") "

                        gDB.amendData(updateSQL, gConn, transaction, paP)

                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "DELETE from WMS_ALERT WHERE SYS_PK=" & paP.AP(sysPK)
                        gDB.amendData(updateSQL, gConn, transaction, paP)

                        updateCount += 1
                    End If
                Next

                If updateCount > 0 Then
                    'uiFun.displayMsgNew(MainUDP, "", "READ!", Session("gLang"))
                    transaction.Commit()
                    successFlag = True
                End If
                
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

            If successFlag Then
                BindGV()
            End If

        End If
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        BindGV()
    End Sub

    Protected Sub SortDT(ByVal sortBy As String)
        Dim tempDT, CopyDT As DataTable
        tempDT = Session("AlertDT")

        If Not tempDT Is Nothing Or tempDT.Rows.Count > 0 Then
            CopyDT = tempDT.Copy
            Dim dv As DataView = CopyDT.DefaultView
            Select Case sortBy
                Case "ALRT_SEND_DATE"
                    dv.Sort = "ALRT_SEND_DATE DESC"
                    tempDT = dv.ToTable
                Case "ALRT_PRIORITY"
                    dv.Sort = "PRIORITY_SEQ ASC, ALRT_SEND_DATE DESC"
                    tempDT = dv.ToTable
            End Select

            GridView1.DataSource = tempDT
            Session("AlertDT") = tempDT
            GridView1.DataBind()

        End If
    End Sub
End Class
