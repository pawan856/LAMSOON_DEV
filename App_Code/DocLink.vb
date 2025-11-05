Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web

Public Class DocLink

    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils

    Private lIMP_CODE As String = ""

    Public Sub genLinkBar(ByRef lb As Table, ByVal STORER_CODE As String, ByVal LN_DOCTYPE As String, ByVal LN_DOC_NO As String, ByVal rootDir As String)

        For i As Integer = 0 To lb.Controls.Count - 1
            lb.Controls(i).Visible = False
        Next

        Dim dt As DataTable
        Dim sqlStr As String

        Dim tr As New TableRow
        Dim tc As New TableCell()
        tc.CssClass = "TITLE"
        tc.Style.Add("width", "20%")

        Dim lba As New Label
        lba.Text = "Link:"
        tc.Controls.Add(lba)

        Dim lb1 As New Label
        lb1.Text = "&nbsp;"
        tc.Controls.Add(lb1)

        Dim prevDocType As String = ""
        Dim prevName As String = ""
        Dim prevDocNo As String = ""
        Dim prevSC As String = ""
        Dim prevRowCount As Integer = 0

        'sqlStr = "select * from WMS_LINK_TX WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' order by LN_DOCTYPE_TO, LN_SEQ"
        sqlStr = "select STORER_CODE, LN_DOCTYPE_TO, LN_NAME, LN_DOC_NO_TO from WMS_LINK_TX WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' AND LN_DOCTYPE = '" & gU.dbEncode(LN_DOCTYPE) & "' AND LN_DOC_NO = '" & gU.dbEncode(LN_DOC_NO) & "' order by 1, 2,3,4"
        dt = gDB.getDataTable(sqlStr)
        For i As Integer = 0 To dt.Rows.Count - 1
            If prevDocType <> dt.Rows(i)("LN_DOCTYPE_TO").ToString Then
                If prevRowCount > 1 Then
                    Dim hl As New HtmlAnchor
                    hl.InnerText = prevName
                    hl.HRef = "#"
                    'hl.Attributes.Add("onclick", "document.forms[0].LN_DOCTYPE_TO.value='" & prevDocType & "';window.open('" & rootDir & "cms_search.aspx?menu_code=LOOKUP_LINK&pForm=forms[0]&pFunc=gotoUrl()&pItemList=LN_DOC_NO_TO|&sc=" & STORER_CODE & "&fdt=" & LN_DOCTYPE & "&fdn=" & LN_DOC_NO & "&ar=Y');")
                    hl.Attributes.Add("onclick", "selectLink('" & prevDocType & "', '" & STORER_CODE & "', '" & LN_DOC_NO & "')")
                    tc.Controls.Add(hl)
                Else
                    Dim hl As New HyperLink
                    hl.Text = prevName
                    Select Case prevDocType
                        Case "GR"
                            hl.NavigateUrl = rootDir & "INBOUND/GR/GRMain.aspx?storer_code=" & prevSC & "&gr_code=" & prevDocNo
                        Case "RO"
                            hl.NavigateUrl = rootDir & "INBOUND/RO/ROMain.aspx?storer_code=" & prevSC & "&ro_code=" & prevDocNo
                        Case "CO"
                            hl.NavigateUrl = rootDir & "OUTBOUND/CO/COMain.aspx?storer_code=" & prevSC & "&co_code=" & prevDocNo
                        Case "DO"
                            hl.NavigateUrl = rootDir & "OUTBOUND/DO/DOMain.aspx?storer_code=" & prevSC & "&do_code=" & prevDocNo
                    End Select
                    tc.Controls.Add(hl)
                End If

                Dim lb2 As New Label
                lb2.Text = "&nbsp;"
                tc.Controls.Add(lb2)
                prevRowCount = 0
            End If
            prevDocType = dt.Rows(i)("LN_DOCTYPE_TO").ToString
            prevDocNo = dt.Rows(i)("LN_DOC_NO_TO").ToString
            prevName = dt.Rows(i)("LN_NAME").ToString

            If prevName = "RO" Then
                prevName = "PO"
            ElseIf prevName = "CO" Then
                prevName = "SIR"
            ElseIf prevName = "DO" Then
                prevName = "WIT"
            End If

            prevSC = dt.Rows(i)("STORER_CODE").ToString
            prevRowCount = prevRowCount + 1
        Next

        If prevDocType <> "" Then
            If prevRowCount > 1 Then
                Dim hl As New HtmlAnchor
                hl.InnerText = prevName
                hl.HRef = "#"
                'hl.Attributes.Add("onclick", "document.forms[0].LN_DOCTYPE_TO.value='" & prevDocType & "';window.open('" & rootDir & "cms_search.aspx?menu_code=LOOKUP_LINK&pForm=forms[0]&pFunc=gotoUrl()&pItemList=LN_DOC_NO_TO|&sc=" & STORER_CODE & "&fdt=" & LN_DOCTYPE & "&fdn=" & LN_DOC_NO & "&ar=Y');")
                hl.Attributes.Add("onclick", "selectLink('" & prevDocType & "', '" & STORER_CODE & "','" & LN_DOC_NO & "')")
                tc.Controls.Add(hl)
            Else
                Dim hl As New HyperLink
                hl.Text = prevName
                Select Case prevDocType
                    Case "GR"
                        hl.NavigateUrl = rootDir & "INBOUND/GR/GRMain.aspx?storer_code=" & prevSC & "&gr_code=" & prevDocNo
                    Case "RO"
                        hl.NavigateUrl = rootDir & "INBOUND/RO/ROMain.aspx?storer_code=" & prevSC & "&ro_code=" & prevDocNo
                    Case "CO"
                        hl.NavigateUrl = rootDir & "OUTBOUND/CO/COMain.aspx?storer_code=" & prevSC & "&co_code=" & prevDocNo
                    Case "DO"
                        hl.NavigateUrl = rootDir & "OUTBOUND/DO/DOMain.aspx?storer_code=" & prevSC & "&do_code=" & prevDocNo
                End Select
                tc.Controls.Add(hl)
            End If
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
        End If

        Dim tempText As String
        tempText = "<input type='hidden' name='LN_DOC_NO_TO' value=''><input type='hidden' name='LN_DOCTYPE_TO' value=''>" & _
        "<script>function gotoUrl() {" & _
        "if (document.forms[0].LN_DOCTYPE_TO.value == ""GR"") " & _
        "{document.hiddenForm.action='" & rootDir & "INBOUND/GR/GRMain.aspx';setInterfaceDataToForm(document.hiddenForm, ""storer_code"", """ & STORER_CODE & """);setInterfaceDataToForm(document.hiddenForm, ""gr_code"", document.forms[0].LN_DOC_NO_TO.value);} " & _
        "else if (document.forms[0].LN_DOCTYPE_TO.value == ""RO"") " & _
        "{document.hiddenForm.action='" & rootDir & "INBOUND/RO/ROMain.aspx';setInterfaceDataToForm(document.hiddenForm, ""storer_code"", """ & STORER_CODE & """);setInterfaceDataToForm(document.hiddenForm, ""ro_code"", document.forms[0].LN_DOC_NO_TO.value);} " & _
        "else if (document.forms[0].LN_DOCTYPE_TO.value == ""DO"") " & _
        "{document.hiddenForm.action='" & rootDir & "OUTBOUND/DO/DOMain.aspx';setInterfaceDataToForm(document.hiddenForm, ""storer_code"", """ & STORER_CODE & """);setInterfaceDataToForm(document.hiddenForm, ""do_code"", document.forms[0].LN_DOC_NO_TO.value);} " & _
        "else if (document.forms[0].LN_DOCTYPE_TO.value == ""CO"") " & _
        "{document.hiddenForm.action='" & rootDir & "OUTBOUND/CO/COMain.aspx';setInterfaceDataToForm(document.hiddenForm, ""storer_code"", """ & STORER_CODE & """);setInterfaceDataToForm(document.hiddenForm, ""co_code"", document.forms[0].LN_DOC_NO_TO.value);} " & _
        "document.hiddenForm.target='main';document.hiddenForm.submit();}</script>"

        tempText = tempText & "" & _
        "<script>function selectLink(DocType, STORER_CODE, DOC_NO) {" & _
        "removeAllElementFromForm(document.hiddenForm);" & _
        "window.open("""", ""linkLookUp"", ""titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15"");" & _
        "document.forms[0].LN_DOCTYPE_TO.value = DocType;" & _
        "setInterfaceDataToForm(document.hiddenForm, ""menu_code"", ""LOOKUP_LINK"");" & _
        "setInterfaceDataToForm(document.hiddenForm, ""pForm"", ""forms[0]"");" & _
        "setInterfaceDataToForm(document.hiddenForm, ""pFunc"", ""gotoUrl()"");" & _
        "setInterfaceDataToForm(document.hiddenForm, ""sc"", STORER_CODE);" & _
        "setInterfaceDataToForm(document.hiddenForm, ""fdt"", """ & LN_DOCTYPE & """);" & _
        "setInterfaceDataToForm(document.hiddenForm, ""fdn"", DOC_NO);" & _
        "setInterfaceDataToForm(document.hiddenForm, ""ar"", ""Y"");" & _
        "setInterfaceDataToForm(document.hiddenForm, ""pItemList"", ""LN_DOC_NO_TO|"");" & _
        "document.hiddenForm.action = """ & rootDir & "cms_search.aspx"";" & _
        "document.hiddenForm.target = ""linkLookUp"";" & _
        "document.hiddenForm.submit();" & _
        "}</script>"

        Dim nSpan As New Label
        nSpan.Text = tempText
        tc.Controls.Add(nSpan)
        tc.HorizontalAlign = HorizontalAlign.Right
        tr.Cells.Add(tc)
        lb.Rows.Add(tr)
    End Sub

    Public Function genDocLink(ByVal fromDocType As String, ByVal fromDocNo As String, ByVal toDocType As String, ByVal toDocNo As String, ByVal STORER_CODE As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        REM Create From Link
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim LN_SEQ As Integer
        sqlString = "SELECT max(LN_SEQ) as LN_SEQ FROM WMS_LINK_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
        "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' " & _
        "AND LN_DOCTYPE = '" & gU.dbEncode(fromDocType) & "' " & _
        "AND LN_DOC_NO = '" & gU.dbEncode(fromDocNo) & "' "

        dt = gDB.getDataTable(sqlString, pConn, pTransaction)

        If String.IsNullOrEmpty(dt.Rows(0).Item(0).ToString) Then
            LN_SEQ = 0
        Else
            LN_SEQ = dt.Rows(0).Item(0)
        End If
        dt = Nothing

        sqlString = "INSERT INTO WMS_LINK_TX (" & _
        "IMP_CODE, " & _
        "STORER_CODE, " & _
        "LN_DOCTYPE, " & _
        "LN_DOC_NO, " & _
        "LN_SEQ, " & _
        "LN_DOCTYPE_TO, " & _
        "LN_DOC_NO_TO, " & _
        "LN_DATE, " & _
        "LN_NAME, " & _
        "LN_DESC, " & _
        "LN_REM, " & _
        "SYS_CB, " & _
        "SYS_CD, " & _
        "SYS_LUB, " & _
        "SYS_LUD" & _
        ") VALUES (" & _
        "'" & gU.dbEncode(lIMP_CODE) & "'," & _
        "'" & gU.dbEncode(STORER_CODE) & "'," & _
        "'" & gU.dbEncode(fromDocType) & "'," & _
        "'" & gU.dbEncode(fromDocNo) & "'," & _
        "" & LN_SEQ + 1 & "," & _
        "'" & gU.dbEncode(toDocType) & "'," & _
        "'" & gU.dbEncode(toDocNo) & "'," & _
        "Getdate()," & _
        "'" & gU.dbEncode(toDocType) & "'," & _
        "'" & gU.dbEncode(toDocType) & "'," & _
        "NULL," & _
        "'" & Session("usr_id") & "', " & _
        "Getdate(), " & _
        "'" & Session("usr_id") & "'," & _
        "Getdate())"

        If pConn Is Nothing Then
            Call gDB.amendData(sqlString)
        Else
            If pTransaction Is Nothing Then
                Call gDB.amendData(sqlString, pConn)
            Else
                Call gDB.amendData(sqlString, pConn, pTransaction)
            End If
        End If


        sqlString = "SELECT max(LN_SEQ) as LN_SEQ FROM WMS_LINK_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
        "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' " & _
        "AND LN_DOCTYPE = '" & gU.dbEncode(toDocType) & "' " & _
        "AND LN_DOC_NO = '" & gU.dbEncode(toDocNo) & "' "

        dt = gDB.getDataTable(sqlString, pConn, pTransaction)


        If String.IsNullOrEmpty(dt.Rows(0).Item(0).ToString) Then
            LN_SEQ = 0
        Else
            LN_SEQ = dt.Rows(0).Item(0)
        End If
        dt = Nothing

        REM Create To Link
        sqlString = "INSERT INTO WMS_LINK_TX (" & _
        "IMP_CODE, " & _
        "STORER_CODE, " & _
        "LN_DOCTYPE, " & _
        "LN_DOC_NO, " & _
        "LN_SEQ, " & _
        "LN_DOCTYPE_TO, " & _
        "LN_DOC_NO_TO, " & _
        "LN_DATE, " & _
        "LN_NAME, " & _
        "LN_DESC, " & _
        "LN_REM, " & _
        "SYS_CB, " & _
        "SYS_CD, " & _
        "SYS_LUB, " & _
        "SYS_LUD" & _
        ") VALUES (" & _
        "'" & gU.dbEncode(lIMP_CODE) & "'," & _
        "'" & gU.dbEncode(STORER_CODE) & "'," & _
        "'" & gU.dbEncode(toDocType) & "'," & _
        "'" & gU.dbEncode(toDocNo) & "'," & _
        "" & LN_SEQ + 1 & "," & _
        "'" & gU.dbEncode(fromDocType) & "'," & _
        "'" & gU.dbEncode(fromDocNo) & "'," & _
        "Getdate()," & _
        "'" & gU.dbEncode(fromDocType) & "'," & _
        "'" & gU.dbEncode(fromDocType) & "'," & _
        "NULL," & _
        "'" & Session("usr_id") & "', " & _
        "Getdate(), " & _
        "'" & Session("usr_id") & "'," & _
        "Getdate())"

        If pConn Is Nothing Then
            Call gDB.amendData(sqlString)
        Else
            If pTransaction Is Nothing Then
                Call gDB.amendData(sqlString, pConn)
            Else
                Call gDB.amendData(sqlString, pConn, pTransaction)
            End If
        End If

        genDocLink = True
    End Function

    Public Sub New()
        Dim sqlString As String = "select IMP_CODE from WMS_SETTINGS_APPLICATION"
        Dim dt As New DataTable
        dt = gDB.getDataTable(sqlString)
        lIMP_CODE = dt.Rows(0).Item(0).ToString
        dt = Nothing
    End Sub
End Class
