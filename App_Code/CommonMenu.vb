Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.UI.WebControls

Public Class CommonMenu
    Inherits System.Web.UI.Page

    Private lHaveCheckList As String = "N"
    Private lHaveAttachments As String = "N"
    Private lHaveNotes As String = "N"
    Private lHaveTasks As String = "N"
    Private lHaveEmail As String = "N"
    Private lHaveHistory As String = "N"
    Private ldocNo As String = ""
    Private ldocType As String = ""
    Private ldocHeader As String = ""
    Private lParentDir As String = "../"

    Private gu As New GeneralUtils
    Private gDB As New GlobalDBFunc

    Public Property parentDir() As String
        Get
            Return lParentDir
        End Get
        Set(ByVal Value As String)
            lParentDir = Value
        End Set
    End Property

    Public Property docHeader() As String
        Get
            Return ldocHeader
        End Get
        Set(ByVal Value As String)
            ldocHeader = Value
        End Set
    End Property

    Public Property docNo() As String
        Get
            Return ldocType
        End Get
        Set(ByVal Value As String)
            ldocNo = Value
        End Set
    End Property

    Public Property docType() As String
        Get
            Return ldocType
        End Get
        Set(ByVal Value As String)
            ldocType = Value
        End Set
    End Property

    Public Property haveCheckList() As String
        Get
            Return lHaveCheckList
        End Get
        Set(ByVal Value As String)
            lHaveCheckList = Value
        End Set
    End Property

    Public Property haveAttachments() As String
        Get
            Return lHaveAttachments
        End Get
        Set(ByVal Value As String)
            lHaveAttachments = Value
        End Set
    End Property

    Public Property haveNotes() As String
        Get
            Return lHaveNotes
        End Get
        Set(ByVal Value As String)
            lHaveNotes = Value
        End Set
    End Property

    Public Property haveTasks() As String
        Get
            Return lHaveTasks
        End Get
        Set(ByVal Value As String)
            lHaveTasks = Value
        End Set
    End Property

    Public Property haveEmail() As String
        Get
            Return lHaveEmail
        End Get
        Set(ByVal Value As String)
            lHaveEmail = Value
        End Set
    End Property

    Public Property haveHistory() As String
        Get
            Return lHaveHistory
        End Get
        Set(ByVal Value As String)
            lHaveHistory = Value
        End Set
    End Property

    Public Sub genCM(ByRef lb As Table)
        Dim tr As New TableRow
        Dim tc As New TableCell()
        tc.CssClass = "TITLE"
        tc.Wrap = False
        tc.Style.Add("width", "auto")

        Dim lb1 As New Label
        lb1.Text = "<b>" & ldocHeader & "</b>"
        tc.Controls.Add(lb1)
        tr.Cells.Add(tc)

        tc = New TableCell()
        tc.CssClass = "TITLE"
        tc.Style.Add("width", "100%")

        lb1 = New Label
        lb1.Text = "&nbsp;"
        tc.Controls.Add(lb1)

        If haveCheckList = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "Check Lists"
            lbutton.Attributes.Add("onclick", "goToCheckList();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToCheckList() {window.open('" & lParentDir & "checkList.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        If haveAttachments = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "Attachments"
            lbutton.Attributes.Add("onclick", "goToAttach();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToAttach() {window.open('" & lParentDir & "attachments.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        If haveNotes = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "Notes"
            lbutton.Attributes.Add("onclick", "goToNotes();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToNotes() {window.open('" & lParentDir & "note.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        If haveTasks = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "Tasks"
            lbutton.Attributes.Add("onclick", "goToTasks();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToTasks() {window.open('" & lParentDir & "task.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        If haveEmail = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "Email"
            lbutton.Attributes.Add("onclick", "goToEmail();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToEmail() {window.open('" & lParentDir & "email.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        If haveHistory = "Y" Then
            Dim lbutton As New HtmlInputButton
            lbutton.Value = "History"
            lbutton.Attributes.Add("onclick", "goToHistory();")
            lbutton.Attributes.Add("Class", "all_button")
            lbutton.Disabled = True
            tc.Controls.Add(lbutton)
            Dim lb2 As New Label
            lb2.Text = "&nbsp;"
            tc.Controls.Add(lb2)
            Dim nSpan As New Label
            nSpan.Text = "<script>function goToHistory() {window.open('" & lParentDir & "history.aspx');}</script>"
            tc.Controls.Add(nSpan)
        End If

        tc.HorizontalAlign = HorizontalAlign.Right

        tr.Cells.Add(tc)
        lb.Rows.Add(tr)
    End Sub

    Public Sub New(Optional ByVal docType As String = Nothing, Optional ByVal docHeader As String = Nothing, Optional ByVal docNo As String = Nothing)
        ldocHeader = docHeader
        ldocType = docType
        ldocNo = docNo

        lHaveCheckList = "N"
        lHaveAttachments = "N"
        lHaveNotes = "N"
        lHaveTasks = "N"
        lHaveEmail = "N"
        lHaveHistory = "N"
    End Sub
End Class
