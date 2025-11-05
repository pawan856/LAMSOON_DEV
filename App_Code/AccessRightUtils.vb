Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.UI.WebControls
Imports AjaxControlToolkit

Public Class AccessRightUtils
    Inherits System.Web.UI.Page

    Private lsec_read As String = "N"
    Private lsec_write As String = "N"
    Private lsec_report As String = "N"
    Private lsec_print As String = "N"
    Private lsec_access As String = "N"
    Private lsessionExpired As String = "N"

    Private lsec_viewMode As String = "N"

    Private buttonList As New List(Of String)

    Private gu As New GeneralUtils
    Private gDB As New GlobalDBFunc

    Private ServerIP As String = System.Configuration.ConfigurationManager.AppSettings.Item("WEB_SERVER_IP")

    Public Delegate Function CustomizeControl(ByVal ctrl As Control) As Boolean
    Public Delegate Function CustomizeGridViewControl(ByVal ctrl As Control, ByRef AddCtrltoCell As ArrayList) As Boolean

    Private Structure col_def_type
        Dim DBField As String
        Dim gvField As String
        Dim colSeq As Integer
        Dim fieldLabelName As String
    End Structure

    Private col_def As New List(Of col_def_type)

    Public Property sessionExpired() As String
        Get
            Return lsessionExpired
        End Get
        Set(ByVal Value As String)
            lsessionExpired = Value
        End Set
    End Property

    Public Property sec_read() As String
        Get
            Return lsec_read
        End Get
        Set(ByVal Value As String)
            lsec_read = Value
        End Set
    End Property

    Public Property sec_write() As String
        Get
            Return lsec_write
        End Get
        Set(ByVal Value As String)
            lsec_write = Value
        End Set
    End Property

    Public Property sec_report() As String
        Get
            Return lsec_report
        End Get
        Set(ByVal Value As String)
            lsec_report = Value
        End Set
    End Property

    Public Property sec_print() As String
        Get
            Return lsec_print
        End Get
        Set(ByVal Value As String)
            lsec_print = Value
        End Set
    End Property

    Public Property sec_viewMode() As String
        Get
            Return lsec_viewMode
        End Get
        Set(ByVal Value As String)
            lsec_viewMode = Value
        End Set
    End Property


    Public Sub Force_PageEndCtrlClear(ByRef p As Page, Optional ByVal ClearCtrl As Boolean = True)
        Dim ServerIP As String = HttpContext.Current.Request.ServerVariables("HTTP_HOST") & "/" & Split(HttpContext.Current.Request.ServerVariables("URL"), "/")(1).ToString
        If ClearCtrl Then p.Controls.Clear()

        HttpContext.Current.Session.RemoveAll()
        HttpContext.Current.Session.Abandon()
        FormsAuthentication.SignOut()

        HttpContext.Current.Response.Clear()
        HttpContext.Current.Response.Write("<scr" & _
                                                "ipt>alert('Server Update. Please login again.');top.location.href='./';</scr" & "ipt>")
        HttpContext.Current.Response.End()
    End Sub

    Public Sub checkSessionExpired(ByRef aspxPage As System.Web.UI.Page)
        Dim ServerIP As String = HttpContext.Current.Request.ServerVariables("HTTP_HOST") & "/" & Split(HttpContext.Current.Request.ServerVariables("URL"), "/")(1).ToString
        If lsessionExpired = "Y" Then
            Dim client1 As ClientScriptManager
            Dim strScript As String
            client1 = aspxPage.ClientScript

            strScript = "alert('Server Update. Please login again.');parent.location.href='./';"
            If (Not client1.IsStartupScriptRegistered(aspxPage.GetType(), "0000")) Then
                client1.RegisterStartupScript(aspxPage.GetType(), "0000", strScript, True)
            End If
        End If
    End Sub

    Public Sub addColDef(ByVal DBField As String, ByVal gvField As String, ByRef colseq As Integer, Optional ByVal fieldLabelName As String = "", Optional ByVal autoIncreament As Boolean = True)
        Dim tcoldef As New col_def_type
        tcoldef.DBField = DBField.ToString.ToUpper
        tcoldef.gvField = gvField.ToString.ToUpper
        tcoldef.colSeq = colseq

        If autoIncreament Then
            colseq += 1
        End If

        tcoldef.fieldLabelName = fieldLabelName.ToString.ToUpper
        col_def.Add(tcoldef)
    End Sub

    Private Sub hideGVControl(ByRef ctl As Control, ByVal DD_DISP As String, ByVal DD_DISABLE As String, ByVal DD_READONLY As String, ByVal DD_FONTTYPE As String, ByVal DD_BACKCOLOR As String)
        If DD_DISP = "N" Then
            ctl.Visible = False
        End If

        If DD_DISABLE = "Y" Then
            If TypeOf ctl Is TextBox Then
                CType(ctl, TextBox).Enabled = False
            ElseIf TypeOf ctl Is CheckBox Then
                CType(ctl, CheckBox).Enabled = False
            ElseIf TypeOf ctl Is DropDownList Then
                CType(ctl, DropDownList).Enabled = False
            ElseIf TypeOf ctl Is Button Then
                CType(ctl, Button).Enabled = False
            ElseIf TypeOf ctl Is GridView Then
                CType(ctl, GridView).Enabled = False
            ElseIf TypeOf ctl Is LinkButton Then
                CType(ctl, LinkButton).Enabled = False
            ElseIf TypeOf ctl Is HyperLink Then
                CType(ctl, HyperLink).Enabled = False
            ElseIf TypeOf ctl Is ImageButton Then
                CType(ctl, ImageButton).Enabled = False
            End If
        End If
        If DD_READONLY = "Y" Then
            If TypeOf ctl Is TextBox Then
                CType(ctl, TextBox).ReadOnly = True
                CType(ctl, TextBox).BorderWidth = 0
                CType(ctl, TextBox).BackColor = Drawing.Color.Transparent

                Dim tempLabel As New Label
                tempLabel.ID = "DL_" & CType(ctl, TextBox).ID
                'tempLabel.Width = CType(ctl, TextBox).Width
                tempLabel.Font.Size = CType(ctl, TextBox).Font.Size

                tempLabel.Text = CType(ctl, TextBox).Text & "&nbsp;"

                If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                    ctl.Parent.Controls.Add(tempLabel)
                    CType(ctl, TextBox).Visible = False
                End If

            ElseIf TypeOf ctl Is CheckBox Then
                CType(ctl, CheckBox).Enabled = False
            ElseIf TypeOf ctl Is DropDownList Then
                CType(ctl, DropDownList).Enabled = False

                Dim tempLabel As New Label
                tempLabel.ID = "DL_" & CType(ctl, DropDownList).ID
                tempLabel.Style.Add("width", "auto")
                tempLabel.Style.Add("font-size", "11px")
                If CType(ctl, DropDownList).SelectedValue <> "" Then
                    tempLabel.Text = CType(ctl, DropDownList).SelectedItem.Text
                Else
                    tempLabel.Text = ""
                End If

                If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                    ctl.Parent.Controls.Add(tempLabel)
                    CType(ctl, DropDownList).Visible = False
                End If

            ElseIf TypeOf ctl Is Button Then
                CType(ctl, Button).Enabled = False
            ElseIf TypeOf ctl Is GridView Then
                CType(ctl, GridView).Enabled = False
            ElseIf TypeOf ctl Is LinkButton Then
                CType(ctl, LinkButton).Enabled = False
            ElseIf TypeOf ctl Is HyperLink Then
                CType(ctl, HyperLink).Enabled = False
            ElseIf TypeOf ctl Is ImageButton Then
                CType(ctl, ImageButton).Enabled = False
            End If
        End If
        If DD_FONTTYPE = "BOLD" Then
            If TypeOf ctl Is TextBox Then
                CType(ctl, TextBox).Font.Bold = True
            ElseIf TypeOf ctl Is CheckBox Then
                CType(ctl, CheckBox).Font.Bold = True
            ElseIf TypeOf ctl Is DropDownList Then
                If DD_READONLY = "Y" Then
                    Dim tempLabel As New Label
                    tempLabel = CType(ctl.Parent.FindControl("DL_" & CType(ctl, DropDownList).ID), Label)
                    tempLabel.Font.Bold = True
                Else
                    CType(ctl, DropDownList).Font.Bold = True
                End If

            ElseIf TypeOf ctl Is Button Then
                CType(ctl, Button).Font.Bold = True
            ElseIf TypeOf ctl Is HyperLink Then
                CType(ctl, HyperLink).Font.Bold = True
            End If
        ElseIf DD_FONTTYPE = "ITALIC" Then
            If TypeOf ctl Is TextBox Then
                CType(ctl, TextBox).Font.Italic = True
            ElseIf TypeOf ctl Is CheckBox Then
                CType(ctl, CheckBox).Font.Italic = True
            ElseIf TypeOf ctl Is DropDownList Then
                If DD_READONLY = "Y" Then
                    Dim tempLabel As New Label
                    tempLabel = CType(ctl.Parent.FindControl("DL_" & CType(ctl, DropDownList).ID), Label)
                    tempLabel.Font.Italic = True
                Else
                    CType(ctl, DropDownList).Font.Italic = True
                End If

            ElseIf TypeOf ctl Is Button Then
                CType(ctl, Button).Font.Italic = True
            ElseIf TypeOf ctl Is HyperLink Then
                CType(ctl, HyperLink).Font.Italic = True
            End If
        End If
        If DD_BACKCOLOR <> "" Then
            If TypeOf ctl Is TextBox Then
                'CType(ctl, TextBox).BackColor = 
                CType(ctl, TextBox).Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
            ElseIf TypeOf ctl Is CheckBox Then
                CType(ctl, CheckBox).Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
            ElseIf TypeOf ctl Is DropDownList Then
                If DD_READONLY = "Y" Then
                    Dim tempLabel As New Label
                    tempLabel = CType(ctl.Parent.FindControl("DL_" & CType(ctl, DropDownList).ID), Label)

                    tempLabel.Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
                Else
                    CType(ctl, DropDownList).Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
                End If
            ElseIf TypeOf ctl Is Button Then
                CType(ctl, Button).Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
            ElseIf TypeOf ctl Is HyperLink Then
                CType(ctl, HyperLink).Style.Add("BACKGROUND-COLOR", "#" & DD_BACKCOLOR)
            End If
        End If
    End Sub

    Public Sub hideGVForStorer(ByRef gv As GridView, ByVal Storer_code As String, ByVal DD_TYPE As String, ByVal DD_DBTABLE As String, Optional ByRef e As GridViewRowEventArgs = Nothing)
        Dim sqlString As String = ""
        Dim dt As New DataTable

        sqlString = "SELECT DD_DBFIELD,DD_DISP, DD_DISABLE,DD_READONLY, DD_FONTTYPE, DD_BACKCOLOR from WMS_DETDISP_CONTROL where IMP_CODE = '" & gu.dbEncode(Session("imp_code")) & "' " & _
        "and STORER_CODE = '" & gu.dbEncode(Storer_code) & "' " & _
        "and DD_TYPE = '" & gu.dbEncode(DD_TYPE) & "' " & _
        "and DD_DBTABLE = '" & gu.dbEncode(DD_DBTABLE) & "' "
        dt = gDB.getDataTable(sqlString)
        For iCount As Integer = 0 To dt.Rows.Count - 1
            For i As Integer = 0 To col_def.Count - 1
                If col_def(i).DBField.ToString.ToUpper = dt.Rows(iCount)("DD_DBFIELD").ToString.ToUpper Then
                    If Not IsNothing(e) Then
                        Dim lctl = e.Row.FindControl(col_def(i).gvField)

                        If lctl IsNot Nothing Then
                            hideGVControl(lctl, dt.Rows(iCount)("DD_DISP").ToString, dt.Rows(iCount)("DD_DISABLE").ToString, dt.Rows(iCount)("DD_READONLY").ToString, dt.Rows(iCount)("DD_FONTTYPE").ToString, dt.Rows(iCount)("DD_BACKCOLOR").ToString)
                            If dt.Rows(iCount)("DD_DISP").ToString = "N" Then
                                If col_def(i).fieldLabelName <> "" Then
                                    e.Row.FindControl(col_def(i).fieldLabelName).Visible = False
                                Else
                                    gv.Columns(col_def(i).colSeq).Visible = False
                                    CType(gv.Controls(0).Controls(0), GridViewRow).Cells(col_def(i).colSeq).Visible = False
                                End If
                            End If
                        End If
                    Else
                        For gvi As Integer = 0 To gv.Rows.Count - 1
                            Dim lctl = gv.Rows(gvi).FindControl(col_def(i).gvField)
                            If lctl IsNot Nothing Then
                                hideGVControl(lctl, dt.Rows(iCount)("DD_DISP").ToString, dt.Rows(iCount)("DD_DISABLE").ToString, dt.Rows(iCount)("DD_READONLY").ToString, dt.Rows(iCount)("DD_FONTTYPE").ToString, dt.Rows(iCount)("DD_BACKCOLOR").ToString)
                                If dt.Rows(iCount)("DD_DISP").ToString = "N" Then
                                    If col_def(i).fieldLabelName <> "" Then
                                        gv.Rows(gvi).FindControl(col_def(i).fieldLabelName).Visible = False
                                    Else
                                        gv.Columns(col_def(i).colSeq).Visible = False
                                        CType(gv.Controls(0).Controls(0), GridViewRow).Cells(col_def(i).colSeq).Visible = False
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Next
        dt.Dispose()
    End Sub

    Public Sub hideGVRow(ByRef gv As GridView, ByVal gvr As System.Web.UI.WebControls.GridViewRow)
        Dim rowCell As TableCellCollection = gvr.Cells
        For Each itemCell As TableCell In rowCell
            For Each ctl As Control In itemCell.Controls
                If TypeOf ctl Is TextBox Then
                    CType(ctl, TextBox).Enabled = False
                ElseIf TypeOf ctl Is CheckBox Then
                    CType(ctl, CheckBox).Enabled = False
                ElseIf TypeOf ctl Is DropDownList Then
                    CType(ctl, DropDownList).Enabled = False
                ElseIf TypeOf ctl Is Button Then
                    CType(ctl, Button).Enabled = False
                ElseIf TypeOf ctl Is GridView Then
                    CType(ctl, GridView).Enabled = False
                ElseIf TypeOf ctl Is LinkButton Then
                    CType(ctl, LinkButton).Enabled = False
                ElseIf TypeOf ctl Is HyperLink Then
                    CType(ctl, HyperLink).Enabled = False
                ElseIf TypeOf ctl Is ImageButton Then
                    CType(ctl, ImageButton).Enabled = False
                ElseIf TypeOf ctl Is Image Then
                    CType(ctl, Image).Visible = False
                End If
            Next
        Next
        gvr.CssClass = "GRAY"
        gvr.CssClass = ""
    End Sub

    Public Sub New()
        lsec_read = "N"
        lsec_write = "N"
        lsec_report = "N"
        lsec_print = "N"
        lsec_access = "N"
        lsessionExpired = "N"
        lsec_viewMode = "N"
    End Sub

    Public Sub New(ByVal fun_code As String, ByVal usr_id As String, Optional ByRef aspxPage As System.Web.UI.Page = Nothing)
        If Session("usr_id") = "" Then
            lsessionExpired = "Y"
            Call checkSessionExpired(aspxPage)
        End If

        Dim sqlString As String = ""
        Dim dt As New DataTable

        sqlString = "SELECT sa.sec_read, sa.sec_write, sa.sec_report, sa.sec_print, sa.sec_access " & _
                    "from wms_sec_access sa, wms_user_group_alloc uga " & _
                    "where sa.grp_code = uga.grp_code and fun_code = '" & gu.dbEncode(fun_code) & "' " & _
                    "and uga.usr_id = '" & gu.dbEncode(usr_id) & "'"
        dt = gDB.getDataTable(sqlString)

        If dt.Rows.Count > 0 Then
            lsec_read = "Y"

            For Each row As DataRow In dt.Rows
                If lsec_write = "" Then
                    lsec_write = row.Item("sec_write").ToString
                Else
                    If row.Item("sec_write").ToString = "Y" AndAlso lsec_write = "N" Then
                        lsec_write = row.Item("sec_write").ToString
                    End If
                End If
                
                If lsec_report = "" Then
                    lsec_report = row.Item("sec_report").ToString
                Else
                    If row.Item("sec_report").ToString = "Y" AndAlso lsec_report = "N" Then
                        lsec_report = row.Item("sec_report").ToString
                    End If
                End If

                If lsec_print = "" Then
                    lsec_print = row.Item("sec_print").ToString
                Else
                    If row.Item("sec_print").ToString = "Y" AndAlso lsec_print = "N" Then
                        lsec_print = row.Item("sec_print").ToString
                    End If
                End If

                If lsec_access = "" Then
                    lsec_access = row.Item("sec_access").ToString
                Else
                    If row.Item("sec_access").ToString = "Y" AndAlso lsec_access = "N" Then
                        lsec_access = row.Item("sec_access").ToString
                    End If
                End If                
            Next
        End If

        sqlString = "select sa.fun_code " & _
                    "from wms_function f, wms_sec_access sa, wms_user_group_alloc uga " & _
                    "where sa.grp_code = uga.grp_code " & _
                    "and sa.fun_code = f.fun_code " & _
                    "and f.fun_type = 'BUTTON' " & _
                    "and f.fun_parent_code = '" & gu.dbEncode(fun_code) & "' " & _
                    "and uga.usr_id = '" & gu.dbEncode(usr_id) & "' "

        dt = gDB.getDataTable(sqlString)

        For i = 0 To dt.Rows.Count - 1
            buttonList.Add(dt.Rows(i).Item("fun_code").ToString.Trim)
        Next

        dt.Dispose()
    End Sub

    Public Function hasBtnRight(ByVal btn_fun_code As String) As Boolean
        If buttonList.Contains(btn_fun_code) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function hideForm(ByRef pid As Page, Optional ByRef hiddenFlag As HiddenField = Nothing, _
                             Optional ByVal callBackSub As CustomizeControl = Nothing, _
                             Optional ByVal gvcallBack As CustomizeGridViewControl = Nothing, _
                             Optional ByRef exceptionList As List(Of String) = Nothing) As Boolean
        hideForm = False

        If lsec_access = "N" Then
            hideForm = True
            Call Force_PageEndCtrlClear(pid, True)
        Else
            If lsec_viewMode = "Y" Then
                If Not hiddenFlag Is Nothing Then hiddenFlag.Value = "V"

                For Each ctl As Control In pid.Form.Controls
                    verifyControl(ctl, "V", , callBackSub, gvcallBack, exceptionList)
                Next

            ElseIf lsec_write = "N" Then

                If Not hiddenFlag Is Nothing Then hiddenFlag.Value = "V"

                For Each ctl As Control In pid.Form.Controls
                    verifyControl(ctl, "D", , callBackSub, gvcallBack, exceptionList)
                Next
            Else
                If Not hiddenFlag Is Nothing Then hiddenFlag.Value = "E"
            End If
        End If
    End Function

    Private Sub verifyControl(ByRef ctl As Control, ByVal mode As String, _
                               Optional ByRef ctlArray As ArrayList = Nothing, _
                               Optional ByVal CallBackSub As CustomizeControl = Nothing, _
                               Optional ByVal GVCallBack As CustomizeGridViewControl = Nothing, _
                                Optional ByRef exceptionList As List(Of String) = Nothing)

        If CallBackSub IsNot Nothing Then
            If Not CallBackSub(ctl) Then
                settingControls(ctl, mode, ctlArray, CallBackSub, GVCallBack, exceptionList)
            End If
        Else
            settingControls(ctl, mode, ctlArray, CallBackSub, GVCallBack, exceptionList)
        End If
    End Sub

    Public Sub settingControls(ByRef ctl As Control, ByVal mode As String, _
                               Optional ByRef ctlArray As ArrayList = Nothing, _
                               Optional ByVal CallBackSub As CustomizeControl = Nothing, _
                               Optional ByVal GVCallBack As CustomizeGridViewControl = Nothing, _
                               Optional ByRef exceptionList As List(Of String) = Nothing)

        If exceptionList IsNot Nothing AndAlso exceptionList.Contains(ctl.ID) Then
            Exit Sub
        End If
        Select Case TypeName(ctl).ToUpper
            Case "TEXTBOX"
                If mode = "V" Then
                    Dim tempLabel As New Label
                    tempLabel.ID = "DL_" & CType(ctl, TextBox).ID
                    'tempLabel.Width = CType(ctl, TextBox).Width
                    tempLabel.Font.Size = CType(ctl, TextBox).Font.Size

                    tempLabel.Text = CType(ctl, TextBox).Text & "&nbsp;"

                    CType(ctl, TextBox).Visible = False

                    If ctlArray Is Nothing Then
                        If Not TypeOf ctl.Parent Is HtmlForm Then
                            If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                                ctl.Parent.Controls.Add(tempLabel)
                            End If
                        Else
                            CType(ctl, TextBox).Visible = True
                            CType(ctl, TextBox).ReadOnly = True
                            CType(ctl, TextBox).BorderWidth = 0
                            CType(ctl, TextBox).BackColor = Drawing.Color.Transparent
                        End If
                    Else
                        If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                            ctlArray.Add(tempLabel)
                        End If

                    End If
                ElseIf mode = "D" Then
                    CType(ctl, TextBox).Enabled = False
                End If

            Case "BUTTON"
                CType(ctl, Button).Enabled = False

            Case "IMAGEBUTTON"
                If mode = "V" Then
                    CType(ctl, ImageButton).Visible = False
                ElseIf mode = "D" Then
                    CType(ctl, ImageButton).Enabled = False
                End If

            Case "DROPDOWNLIST"
                If mode = "V" Then
                    Dim tempLabel As New Label
                    tempLabel.ID = "DL_" & CType(ctl, DropDownList).ID
                    tempLabel.Style.Add("width", "auto")
                    tempLabel.Font.Size = CType(ctl, DropDownList).Font.Size

                    If CType(ctl, DropDownList).SelectedValue <> "" Then
                        tempLabel.Text = CType(ctl, DropDownList).SelectedItem.Text
                    Else
                        tempLabel.Text = ""
                    End If

                    CType(ctl, DropDownList).Visible = False

                    If ctlArray Is Nothing Then
                        If Not TypeOf ctl.Parent Is HtmlForm Then
                            If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                                ctl.Parent.Controls.Add(tempLabel)
                            End If
                        Else
                            CType(ctl, DropDownList).Visible = True
                            CType(ctl, DropDownList).Enabled = False
                        End If
                    Else
                        If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                            ctlArray.Add(tempLabel)
                        End If
                    End If

                ElseIf mode = "D" Then
                    CType(ctl, DropDownList).Enabled = False
                End If

            Case "CHECKBOX"
                CType(ctl, CheckBox).Enabled = False
            Case "IMAGE"
                If mode = "V" Then
                    CType(ctl, Image).Visible = False
                End If

            Case "HTMLIMAGE"
                If mode = "V" Then
                    CType(ctl, HtmlImage).Visible = False
                End If

            Case "HTMLANCHOR"
                CType(ctl, HtmlAnchor).Visible = False

            Case "HTMLTABLECELL"
                If ctl.HasControls Then
                    Dim cellctlArray As New ArrayList

                    For Each cCtrl As Control In ctl.Controls

                        If CallBackSub IsNot Nothing Then
                            If Not CallBackSub(cCtrl) Then _
                                verifyControl(cCtrl, mode, cellctlArray, CallBackSub, GVCallBack, exceptionList)
                        Else
                            verifyControl(cCtrl, mode, cellctlArray, CallBackSub, GVCallBack, exceptionList)
                        End If

                    Next

                    If cellctlArray.Count > 0 Then
                        For x As Integer = 0 To cellctlArray.Count - 1
                            If ctl.FindControl(cellctlArray(x).ID) Is Nothing Then
                                ctl.Controls.Add(cellctlArray(x))
                            End If
                        Next
                    End If
                End If

            Case "GRIDVIEW"
                Dim tempGV As GridView = CType(ctl, GridView)
                ctlArray = Nothing


                For Each gvRow As GridViewRow In tempGV.Rows
                    For Each itemCell As TableCell In gvRow.Cells
                        Dim gvctlArray As New ArrayList

                        For Each rowCellCtl As Control In itemCell.Controls

                            If GVCallBack IsNot Nothing Then
                                If Not GVCallBack(rowCellCtl, gvctlArray) Then _
                                    verifyControl(rowCellCtl, mode, gvctlArray, CallBackSub, GVCallBack, exceptionList)
                            Else
                                verifyControl(rowCellCtl, mode, gvctlArray, CallBackSub, GVCallBack, exceptionList)
                            End If
                        Next

                        If gvctlArray.Count > 0 Then
                            For x As Integer = 0 To gvctlArray.Count - 1
                                If itemCell.FindControl(gvctlArray(x).ID) Is Nothing Then
                                    itemCell.Controls.Add(gvctlArray(x))
                                End If
                            Next
                        End If
                    Next
                Next

            Case "PANEL"
                Dim PanelCtlArray As New ArrayList

                For Each pCtl As Control In ctl.Controls
                    PanelCtlArray = New ArrayList
                    settingControls(pCtl, mode, PanelCtlArray, , , exceptionList)

                    If PanelCtlArray IsNot Nothing Then
                        If PanelCtlArray.Count > 0 Then
                            For x As Integer = 0 To PanelCtlArray.Count - 1
                                If TypeOf pCtl.Parent Is HtmlTableCell Or TypeOf pCtl.Parent Is TableCell Then
                                    If pCtl.Parent.FindControl(PanelCtlArray(x).ID) Is Nothing Then
                                        pCtl.Parent.Controls.Add(PanelCtlArray(x))
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

            Case "UPDATEPANEL"
                Dim PanelCtlArray As New ArrayList

                For Each uCtl As Control In CType(ctl, UpdatePanel).ContentTemplateContainer.Controls
                    PanelCtlArray = New ArrayList

                    settingControls(uCtl, mode, PanelCtlArray, , , exceptionList)

                    If PanelCtlArray IsNot Nothing Then
                        If PanelCtlArray.Count > 0 Then
                            For x As Integer = 0 To PanelCtlArray.Count - 1
                                If TypeOf uCtl.Parent Is HtmlTableCell Or TypeOf uCtl.Parent Is TableCell Then
                                    If uCtl.Parent.FindControl(PanelCtlArray(x).ID) Is Nothing Then
                                        uCtl.Parent.Controls.Add(PanelCtlArray(x))
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

            Case "HTMLGENERICCONTROL"
                Dim DivCtlArray As New ArrayList

                For Each dCtl As Control In ctl.Controls
                    DivCtlArray = New ArrayList

                    settingControls(dCtl, mode, DivCtlArray, , , exceptionList)

                    If DivCtlArray IsNot Nothing Then
                        If DivCtlArray.Count > 0 Then
                            For x As Integer = 0 To DivCtlArray.Count - 1
                                If TypeOf dCtl.Parent Is HtmlTableCell Or TypeOf dCtl.Parent Is TableCell Then
                                    dCtl.Parent.Controls.Add(DivCtlArray(x))
                                    If dCtl.Parent.FindControl(DivCtlArray(x).ID) Is Nothing Then
                                        dCtl.Parent.Controls.Add(DivCtlArray(x))
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

            Case "COMBOBOX"
                If mode = "V" Then
                    Dim tempLabel As New Label
                    tempLabel.ID = "DL_" & CType(ctl, AjaxControlToolkit.ComboBox).ID
                    tempLabel.Style.Add("width", "auto")
                    tempLabel.Font.Size = CType(ctl, AjaxControlToolkit.ComboBox).Font.Size

                    If CType(ctl, AjaxControlToolkit.ComboBox).SelectedValue <> "" Then
                        tempLabel.Text = CType(ctl, AjaxControlToolkit.ComboBox).SelectedItem.Text
                    Else
                        tempLabel.Text = ""
                    End If

                    CType(ctl, AjaxControlToolkit.ComboBox).Visible = False

                    If ctlArray Is Nothing Then
                        If Not TypeOf ctl.Parent Is HtmlForm Then
                            If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                                ctl.Parent.Controls.Add(tempLabel)
                            End If
                        Else
                            CType(ctl, AjaxControlToolkit.ComboBox).Visible = True
                            CType(ctl, AjaxControlToolkit.ComboBox).Enabled = False
                        End If
                    Else
                        If ctl.Parent.FindControl(tempLabel.ID) Is Nothing Then
                            ctlArray.Add(tempLabel)
                        End If
                    End If

                ElseIf mode = "D" Then
                    CType(ctl, AjaxControlToolkit.ComboBox).Enabled = False
                End If

            Case "CALENDAREXTENDER"                
                CType(ctl, CalendarExtender).EnabledOnClient = False
        End Select
    End Sub

    Public Shared Sub storeLoginInfo(ByVal userid As String, ByVal userType As String, ByVal usrStorer As String, Optional ByVal gLang As String = "E")
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim lDBFun As New DBfunc

        HttpContext.Current.Session("gLang") = gLang

        If HttpContext.Current.Session("gLang") = "E" Then
            HttpContext.Current.Session("gSelectLabel") = "SELECT"
        Else
            HttpContext.Current.Session("gSelectLabel") = "請選擇"
        End If

        HttpContext.Current.Session("usr_id") = userid
        HttpContext.Current.Session("usr_type") = userType
        HttpContext.Current.Session("usr_pref_storer") = usrStorer


        Dim impsql As String = "select imp_code from wms_settings_application"
        Dim imp_code As String = ""

        imp_code = lDBFun.getValueFromSQL(impsql)

        If imp_code <> "" Then
            HttpContext.Current.Session("IMP_CODE") = imp_code
        Else
            HttpContext.Current.Session("IMP_CODE") = "WMS"
        End If
    End Sub

    Public Sub checkLoginStatus()
        REM*******************************************************************
        REM Applying web.config<appSettings> value into Cache,
        REM The cache entries for each browser would be shared by all pages.

        Dim app_config As New AppConfig()
        app_config.SetCache()
        app_config.Dispose()
        REM*******************************************************************

        'If Membership.GetUser IsNot Nothing Then
        '    If Membership.GetUser.UserName <> "" AndAlso HttpContext.Current.Session("usr_id") = "" Then
        '        Dim SQLString As String = ""
        '        Dim dt As New DataTable

        '        SQLString = "SELECT usr_id, usr_pwd,usr_type,usr_pref_storer from " & HttpContext.Current.Cache("DB_USER") & _
        '                          " where usr_id = '" & Membership.GetUser.UserName & "'"

        '        dt = gDB.getDataTable(SQLString)

        '        If dt.Rows.Count > 0 Then
        '            AccessRightUtils.storeLoginInfo(Membership.GetUser.UserName, dt.Rows.Item(0)("USR_TYPE").ToString, dt.Rows.Item(0)("usr_pref_storer").ToString)
        '        Else
        '            Force_PageEndCtrlClear(Me, True)
        '        End If
        '    ElseIf Membership.GetUser.UserName = "" AndAlso Session("usr_id") = "" Then
        '        Force_PageEndCtrlClear(Me, True)
        '    End If
        'Else
        '    If (HttpContext.Current.Session("usr_id") IsNot Nothing OrElse HttpContext.Current.Session("usr_id") <> "") AndAlso _
        '    (HttpContext.Current.Session("usr_type") IsNot Nothing OrElse HttpContext.Current.Session("usr_type") <> "") Then

        '        storeLoginInfo(HttpContext.Current.Session("usr_id"), HttpContext.Current.Session("usr_type"), HttpContext.Current.Session("usr_pref_storer"))
        '    Else
        '        Force_PageEndCtrlClear(Me, False)
        '    End If
        'End If
    End Sub

    Public Sub setFieldCustomize(ByRef lPage As Page, ByVal fun_code As String, ByVal lstorer_code As String, Optional ByVal tableName As String = "", Optional ByVal tableMode As String = "")
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim fieldDT As DataTable

        Dim ctrlName As String = ""

        Dim tempSQL As String = ""


        paP = New GlobalDBFunc.DBCmdPara

        If tableName <> "" Then
            Select Case tableMode
                Case "M"
                    tempSQL = " AND FLDO_MASTER_TABLE_NAME=" & paP.AP(tableName)
                Case "D"
                    tempSQL = " AND FLDO_TABLE_NAME=" & paP.AP(tableName)
                Case Else
                    tempSQL = " AND FLDO_TABLE_NAME=" & paP.AP(tableName)
            End Select
        Else
            tempSQL = ""
        End If

        selectSQL = "Select FLDO_FIELD_NAME,FLDO_FILED_CTRL, FLDO_FIELD_OPTION,FLDO_MASTER_TABLE_NAME, FLDO_TABLE_NAME from WMS_FIELD_OPTION " & _
                    "Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(lstorer_code) & " AND FUN_CODE=" & paP.AP(fun_code) & _
                    tempSQL

        fieldDT = gDB.getDataTable(selectSQL, , , , paP)

        If fieldDT.Rows.Count > 0 Then
            Dim ctrl As Control
            For i = 0 To fieldDT.Rows.Count - 1
                If fieldDT.Rows(i).Item("FLDO_FIELD_OPTION").ToString.Trim <> "Y" Then
                    ctrlName = fieldDT.Rows(i).Item("FLDO_FILED_CTRL").ToString.Trim
                    ctrl = lPage.Form.FindControl(ctrlName)

                    If Not ctrl Is Nothing Then
                        Select Case TypeName(ctrl).ToUpper
                            Case "TEXTBOX"
                                DirectCast(ctrl, TextBox).Enabled = False
                            Case "BUTTON"
                                DirectCast(ctrl, Button).Enabled = False

                            Case "IMAGEBUTTON"
                                DirectCast(ctrl, ImageButton).Enabled = False

                            Case "DROPDOWNLIST"
                                DirectCast(ctrl, DropDownList).Enabled = False

                            Case "CHECKBOX"
                                DirectCast(ctrl, CheckBox).Enabled = False

                            Case "CHECKBOXLIST"
                                For x = 0 To DirectCast(ctrl, CheckBoxList).Items.Count - 1
                                    DirectCast(ctrl, CheckBoxList).Items(x).Enabled = False
                                Next

                            Case "IMAGE"
                                DirectCast(ctrl, Image).Visible = False

                                'Case "CALENDAREXTENDER"
                                '    DirectCast(ctrl, AjaxControlToolkit.CalendarExtender).Visible = False
                        End Select


                    End If
                Else
                    ctrlName = fieldDT.Rows(i).Item("FLDO_FILED_CTRL").ToString.Trim
                    ctrl = lPage.Form.FindControl(ctrlName)

                    If Not ctrl Is Nothing Then
                        Select Case TypeName(ctrl).ToUpper
                            Case "TEXTBOX"
                                DirectCast(ctrl, TextBox).Enabled = True
                            Case "BUTTON"
                                DirectCast(ctrl, Button).Enabled = True

                            Case "IMAGEBUTTON"
                                DirectCast(ctrl, ImageButton).Enabled = True

                            Case "DROPDOWNLIST"
                                DirectCast(ctrl, DropDownList).Enabled = True

                            Case "CHECKBOX"
                                DirectCast(ctrl, CheckBox).Enabled = True

                            Case "CHECKBOXLIST"
                                For x = 0 To DirectCast(ctrl, CheckBoxList).Items.Count - 1
                                    DirectCast(ctrl, CheckBoxList).Items(x).Enabled = True
                                Next

                            Case "IMAGE"
                                DirectCast(ctrl, Image).Visible = True

                                'Case "CALENDAREXTENDER"
                                '    DirectCast(ctrl, AjaxControlToolkit.CalendarExtender).Visible = True
                        End Select

                    End If
                End If
            Next
        End If


    End Sub

    Public Sub setGVCustomize(ByRef gv As GridView, ByVal fun_code As String, ByVal lstorer_code As String, ByVal tableName As String, ByVal GVcol As String())
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim fieldDT As DataTable

        Dim colName As String = ""

        Dim tempSQL As String = ""


        paP = New GlobalDBFunc.DBCmdPara

        selectSQL = "Select upper(FLDO_FILED_CTRL) as FLDO_FILED_CTRL from WMS_FIELD_OPTION " & _
                    "Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(lstorer_code) & " AND FUN_CODE=" & paP.AP(fun_code) & " AND ISNULL(FLDO_FIELD_OPTION, '') <> 'Y' AND FLDO_FIELD_TYPE is null " & _
                    " AND FLDO_TABLE_NAME=" & paP.AP(tableName)

        Dim tempstr As String = gDB.getCmdSql(selectSQL, paP)
        fieldDT = gDB.getDataTable(selectSQL, , , , paP)

        If fieldDT.Rows.Count > 0 Then
            For i = 0 To GVcol.Length - 1
                If fieldDT.Select("FLDO_FILED_CTRL='" & GVcol(i) & "'").Length > 0 Then
                    gv.Columns(i).Visible = False
                End If
            Next
        End If
    End Sub
End Class
