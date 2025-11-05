Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading

Partial Class DELI_STATUS_DELI_STATUS
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private DB As New DBfunc

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then
            ViewState("co_code") = ""
            ViewState("storer_code") = ""

            ViewState("FTRACK_NO") = Request.Form("ftrack_no")
            ViewState("storer_code") = Request.Form("storer_code")

            Dim webSrv As New AEWebWCF.AEWcfExpressClient
            Dim returnMsg As String = ""

            returnMsg = webSrv.getOrderTraces(ViewState("FTRACK_NO"))

            If Not String.IsNullOrWhiteSpace(returnMsg) Then
                Dim msgLog As PrgmLog = Nothing
                msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "TrackWCF" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                msgLog.writeLog("[Return Message] : " & returnMsg)
            End If

            BindGV()
        End If

    End Sub

    Protected Sub BindGV()
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim tempDT As DataTable

        Dim storer_code As String = ViewState("storer_code")
        Dim FTRACK_NO As String = ViewState("FTRACK_NO")

        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = " SELECT IMP_CODE, STORER_CODE, CO_CODE, DO_CODE, DS_LINE, DS_DATE, DS_REMARKS, " & _
                    " cast(DS_DATE as date) as DS_DATE_D, cast(DS_DATE as time(0)) as DS_TIME, convert(varchar, DS_DATE, 112) as date_sort " & _
                    " FROM WMS_DELIVERY_STATUS " & _
                    " WHERE IMP_CODE=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(storer_code) & " AND DS_FTRACK_NO=" & paP.AP(FTRACK_NO) & _
                    " ORDER by DS_DATE DESC"
        tempDT = gDB.getDataTable(selectSQL, , , , paP)

        Select Case Session("gLang")
            Case "C"
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("zh-HK")
                lblTitle.Text = "運單狀況"
                lblWarn.Text = "很抱歉,沒有運單資料。"
            Case Else
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
                lblTitle.Text = "Delivery Status"
                lblWarn.Text = "Sorry, No tracking record has been found."
        End Select

        Dim preDate As String = ""
        Dim newRow As TableRow
        Dim newCell As TableCell
        Dim newLabel As Label
        Dim newTime As DateTime
        If tempDT.Rows.Count > 0 Then
            lblWarn.Visible = False

            For i = 0 To tempDT.Rows.Count - 1
                If preDate <> tempDT.Rows(i).Item("date_sort").ToString.Trim Then
                    newRow = New TableRow
                    newCell = New TableCell

                    newCell.CssClass = "dateBar"
                    newCell.ColumnSpan = "2"

                    newLabel = New Label

                    If tempDT.Rows(i).Item("DS_DATE_D").ToString.Trim <> "" Then
                        newTime = tempDT.Rows(i).Item("DS_DATE_D")
                        newLabel.Text = newTime.ToLongDateString & " " & newTime.ToString("dddd")
                    Else
                        newLabel.Text = ""
                    End If

                    newCell.Controls.Add(newLabel)
                    newRow.Cells.Add(newCell)

                    statsTable.Rows.Add(newRow)

                    preDate = tempDT.Rows(i).Item("date_sort").ToString.Trim
                End If

                newRow = New TableRow

                newCell = New TableCell
                newLabel = New Label

                newLabel.Text = tempDT.Rows(i).Item("DS_TIME").ToString.Trim
                newCell.Controls.Add(newLabel)
                newCell.HorizontalAlign = HorizontalAlign.Right
                newCell.Width = New Unit("20%")
                newRow.Cells.Add(newCell)

                newCell = New TableCell
                newLabel = New Label

                newLabel.Text = tempDT.Rows(i).Item("DS_REMARKS").ToString.Trim
                newCell.Controls.Add(newLabel)
                newRow.Cells.Add(newCell)


                statsTable.Rows.Add(newRow)
            Next
        Else
            statsTable.Visible = False
            lblWarn.Visible = True
        End If

    End Sub
End Class
