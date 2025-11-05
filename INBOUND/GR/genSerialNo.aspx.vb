Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing

Partial Class INBOUND_GR_genSerialNo
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As New AccessRightUtils
    Private thumb As ThumbGenerator

    Private dt As New DataTable

    Private imp_code As String = "WMS"
    Private storer_code As String = ""
    Private gr_code As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM ****************************
        REM Modify Access Right Here

        'ar = New AccessRightUtils("MAST_IM", Session("usr_id"), Me)
        'ar.hideForm(Me)

        'If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Generate Serial No."
          
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"

            lbl_ImageHd.Text = "Series No. Items"

            SaveBtn1.OnClientClick = "return confirm(""Are you sure to save Serial No. items?"");"
            savebtn2.OnClientClick = "return confirm(""Are you sure to save Serial No. items?"");"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "Generate Serial No."
           
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"

            lbl_ImageHd.Text = "Items"

            saveBtn1.OnClientClick = "return confirm(""确定保存资料?"");"
            saveBtn2.OnClientClick = "return confirm(""确定保存资料?"");"

        End If

        If Not IsPostBack Then
            ViewState("storer_code") = Nothing
            ViewState("gr_code") = Nothing
            ViewState("grd_disp_seq") = Nothing
            ViewState("grd_itm_code") = Nothing
            ViewState("grd_itm_name") = Nothing
            ViewState("grd_rcv_qty") = Nothing

            dt = Nothing
            If Request("sc") IsNot Nothing Then ViewState("storer_code") = Request("sc")
            If Request("gr") IsNot Nothing Then ViewState("gr_code") = Request("gr")
            If Request("i_seq") IsNot Nothing Then ViewState("grd_disp_seq") = Request("i_seq")
            If Request("i_code") IsNot Nothing Then ViewState("grd_itm_code") = Request("i_code")
            If Request("i_name") IsNot Nothing Then ViewState("grd_itm_name") = Request("i_name")
            If Request("i_qty") IsNot Nothing Then ViewState("grd_rcv_qty") = Request("i_qty")
        End If

        storer_code = ViewState("storer_code")
        gr_code = ViewState("gr_code")
        grd_itm_code.Text = ViewState("grd_itm_code")
        grd_itm_name.Text = ViewState("grd_itm_name")
        grd_rcv_qty.Text = ViewState("grd_rcv_qty")

        If Not Session("sn_dt") Is Nothing Then
            dt = Session("sn_dt")

            If dt.Rows.Count > 0 Then
                Call BindGV()
                serial_no.OnClientClick = "return confirm(""This is going to overwrite saved Serial No. items\r\nAre you sure to continue?"");"
            End If
        End If
    End Sub

    Protected Sub BindGV()

        SN_GV.DataSource = dt
        SN_GV.DataBind()

        snHtr.Visible = True
        snSavetr.Visible = True
        snStr.Visible = True
        snGVtr.Visible = True
        SN_GV.Visible = True

    End Sub

    Protected Sub SN_GV_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles SN_GV.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                CType(e.Row.FindControl("GRS_SEQ"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRS_SEQ").ToString.Trim

                CType(e.Row.FindControl("GRS_ITM_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRS_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("GRS_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRS_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("GRS_DATE_IN"), TextBox).Text = cU.chgToYYYYMMDD(DataBinder.Eval(e.Row.DataItem, "GRS_DATE_IN").ToString.Trim)
                CType(e.Row.FindControl("GRS_WARR_STDATE"), TextBox).Text = cU.chgToYYYYMMDD(DataBinder.Eval(e.Row.DataItem, "GRS_WARR_STDATE").ToString.Trim)
                CType(e.Row.FindControl("GRS_WARR_TYPE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRS_WARR_TYPE").ToString.Trim
                CType(e.Row.FindControl("GRS_WARR_PERIOD"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRS_WARR_PERIOD").ToString.Trim
                CType(e.Row.FindControl("GRS_WARR_EXPDATE"), TextBox).Text = cU.chgToYYYYMMDD(DataBinder.Eval(e.Row.DataItem, "GRS_WARR_EXPDATE").ToString.Trim)
                CType(e.Row.FindControl("CUST_NAME"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CUST_NAME").ToString.Trim
               

                CType(e.Row.FindControl("CONS_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CUST_NAME").ToString.Trim
                CType(e.Row.FindControl("CONS_NAME"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CUST_NAME").ToString.Trim
                CType(e.Row.FindControl("CONS_TEL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CUST_NAME").ToString.Trim

                Dim tempLink As HtmlAnchor

                tempLink = CType(e.Row.FindControl("L_GRS_DATE_IN"), HtmlAnchor)
                tempLink.Attributes.Add("onclick", "javascript:pedirFecha(" & CType(e.Row.FindControl("GRS_DATE_IN"), TextBox).ClientID & ",'D_date');")

                tempLink = CType(e.Row.FindControl("L_GRS_WARR_STDATE"), HtmlAnchor)
                tempLink.Attributes.Add("onclick", "javascript:pedirFecha(" & CType(e.Row.FindControl("GRS_WARR_STDATE"), TextBox).ClientID & ",'D_date');")

                tempLink = CType(e.Row.FindControl("L_GRS_WARR_EXPDATE"), HtmlAnchor)
                tempLink.Attributes.Add("onclick", "javascript:pedirFecha(" & CType(e.Row.FindControl("GRS_WARR_EXPDATE"), TextBox).ClientID & ",'D_date');")

                REM **********************
        End Select
    End Sub

    Private Function validateAll() As Boolean

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""

        Dim tmp_itm_flags As String = ""
        Dim tmpImpCode As String = ""

        If imp_code = "" Then tmpImpCode = "WMS" Else tmpImpCode = imp_code

        If validateAll() Then
            Try

                If cU.gfBuildDataTableforGridView(dt, SN_GV, True) Then
                    toBringOver(dt)

                    Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", _
                    "<script language=""JavaScript"">window.open('','_self');window.close();</script>")
                End If

            Catch ex As Exception

                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Private Sub toBringOver(ByRef fromTable As DataTable)
        Dim scriptString As String

        Session("sn_dt") = fromTable
       
        scriptString = "<script language=""JavaScript""> window.opener.document.forms(0).submit(); </script>"

        If Not Page.ClientScript.IsClientScriptBlockRegistered(scriptString) Then
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "script", scriptString)
        End If

    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles savebtn2.Click
        Call save()
    End Sub

    Protected Sub serial_no_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles serial_no.Click
        If sn_prefix.Text <> "" And sn_digit.Text <> "" And gU.decodeEmptyCInt(grd_rcv_qty.Text, 0) > 0 Then

            'If dt Is Nothing Then
            '    Dim shemaSql As String = "select * from WMS_GOODSRCV_D_S"

            '    gDB.getDataTable(shemaSql, , , dt)
            'End If

            If dt.Rows.Count > 0 Then
                dt.Clear()
            End If

            Call genItems()
            Call BindGV()
        Else
            Dim javaStr As String = ""

            If sn_prefix.Text = "" Then
                If Session("gLang") = "E" Then
                    javaStr = "alert('Serial No. Prefix Is Missing!');"
                Else
                    javaStr = "alert('Serial No. Prefix Is Missing!');"
                End If

                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "pMiss", javaStr, True)
                End If
            ElseIf sn_digit.Text = "" Then
                If Session("gLang") = "E" Then
                    javaStr = "alert('Serial No. Start From Is Missing!');"
                Else
                    javaStr = "alert('Serial No. Start From Is Missing!');"
                End If

                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "sMiss", javaStr, True)
                End If
            ElseIf gU.decodeEmptyCInt(grd_rcv_qty.Text, 0) = 0 Then
                If Session("gLang") = "E" Then
                    javaStr = "alert('Zero Received Qty. No Serial No Can be Generated.');"
                Else
                    javaStr = "alert('Zero Received Qty. No Serial No Can be Generated.');"
                End If

                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "sMiss", javaStr, True)
                End If
            Else
                ar.Force_PageEndCtrlClear(Me.Page, True)
            End If
        End If
    End Sub

    Private Sub genItems()
        Dim maxCount As Integer = CInt(grd_rcv_qty.Text)
        Dim currCount As Integer = CInt(sn_digit.Text)

        For i As Integer = 1 To maxCount
            Dim tempDataRow As DataRow = dt.NewRow

            tempDataRow.Item("IMP_CODE") = imp_code
            tempDataRow.Item("STORER_CODE") = storer_code
            tempDataRow.Item("GR_CODE") = gr_code
            tempDataRow.Item("GRS_ITM_CODE") = grd_itm_code.Text
            tempDataRow.Item("GRS_SEQ") = i.ToString
            tempDataRow.Item("GRS_DATE_IN") = Now

            tempDataRow.Item("GRS_SERIAL_NO") = sn_prefix.Text & currCount.ToString.PadLeft(3, "0")

            dt.Rows.Add(tempDataRow)
            dt.AcceptChanges()
            currCount += 1
        Next
    End Sub
End Class
