Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.ServiceProcess

Partial Class OUTBOUND_DO_MemoList
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("usr_id") <> "OTSADMIN" OrElse Session("USR_CMD_IDT") <> "Y" Then
            Session.Abandon()
            FormsAuthentication.SignOut()

            Response.Redirect("~/")
        End If
    End Sub

    Protected Sub btnRoll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRoll.Click
        runCmd(False)
    End Sub

    Protected Sub btnCommit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCommit.Click
        runCmd(True)
    End Sub

    Private Sub runCmd(ByVal isCommitBtn As Boolean)
        Dim cnn As SQLConnection
        Dim transaction As SQLTransaction
        Dim runCmdSql As String = ""
        Dim sqlType, actionType As String
        Dim tableNameArray As String()
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim dt As DataTable
        Dim isCommit As Boolean = isCommitBtn
        Dim affectRow As Long
        Dim resTable As New Table
        Dim tmpTR As TableRow
        Dim tmpTD As TableCell
        Dim tmpGV As GridView
        Dim msgLog As New PrgmLog

        msgLog.skipLog = True

        If tb_cmd.Text.Trim = "" Then
            uiFun.displayMsg(Me, "", "No command to be execute!", Session("gLang"))
            Exit Sub
        End If

        cnn = gDB.getConnection()
        transaction = cnn.BeginTransaction

        Try
            runCmdSql = tb_cmd.Text.Trim

            sqlType = ""

            If (UCase(Left(Trim(runCmdSql), 6)) = "UPDATE") Then
                sqlType = "U"
                actionType = "updated"
            ElseIf (UCase(Left(Trim(runCmdSql), 6)) = "DELETE") Then
                sqlType = "D"
                actionType = "deleted"
            ElseIf (UCase(Left(Trim(runCmdSql), 6)) = "SELECT") Then
                sqlType = "S"
            ElseIf (UCase(Left(Trim(runCmdSql), 6)) = "INSERT") Then
                sqlType = "I"
                actionType = "inserted"
            ElseIf (UCase(Left(Trim(runCmdSql), 6)) = "ALTER ") Then
                sqlType = "SYSTEM"
            ElseIf (UCase(Left(Trim(runCmdSql), 5)) = "DESC ") Then
                sqlType = "DESC"
                runCmdSql = UCase(runCmdSql)
            End If

            Select Case (sqlType)

                Case "S"
                    isCommit = False

                    dt = gDB.getDataTable(runCmdSql, cnn, transaction, , , msgLog)


                    dtdiv.Controls.Clear()

                    resTable.BorderStyle = BorderStyle.None
                    resTable.CellPadding = 1
                    resTable.CellSpacing = 1
                    resTable.HorizontalAlign = HorizontalAlign.Center
                    resTable.Width = Unit.Percentage(95)

                    dtdiv.Controls.Add(resTable)

                    If dt.Rows.Count = 0 Then
                        tmpTR = New TableRow
                        tmpTD = New TableCell
                        tmpTD.Text = "No record found!"
                        tmpTR.Controls.Add(tmpTD)
                        resTable.Controls.Add(tmpTR)
                    Else
                        tmpTR = New TableRow
                        tmpTD = New TableCell
                        tmpTR.Controls.Add(tmpTD)
                        resTable.Controls.Add(tmpTR)

                        tmpGV = genResultGV(dt)

                        tmpGV.DataSource = dt
                        tmpGV.DataBind()

                        tmpTD.Controls.Add(tmpGV)
                    End If

                Case "U", "D", "I"
                    affectRow = gDB.amendData2(runCmdSql, cnn, transaction, , msgLog)

                    dtdiv.Controls.Clear()

                    resTable.BorderStyle = BorderStyle.None
                    resTable.CellPadding = 1
                    resTable.CellSpacing = 1
                    resTable.HorizontalAlign = HorizontalAlign.Center
                    resTable.Width = Unit.Percentage(95)

                    dtdiv.Controls.Add(resTable)

                    tmpTR = New TableRow
                    tmpTD = New TableCell
                    tmpTD.Text = "Number of rows affected: " & affectRow

                    If isCommit Then
                        tmpTD.Text = tmpTD.Text & " - Committed"
                    Else
                        tmpTD.Text = tmpTD.Text & " - Rollbacked"
                    End If

                    tmpTR.Controls.Add(tmpTD)
                    resTable.Controls.Add(tmpTR)

                Case "DESC"
                    isCommit = False

                    tableNameArray = Split(runCmdSql, "DESC")

                    cmdPa = New GlobalDBFunc.DBCmdPara

                    runCmdSql = "SELECT COLUMN_NAME as Name, DECODE(NULLABLE,'Y','','NOT NULL') as NullLable," & _
                                   "decode(data_type || '_' || nvl(to_char(data_precision), '') || '_' || nvl(to_char(data_precision), '0') || '_' || TO_CHAR(data_length), 'NUMBER__0_22', 'INTEGER', " & _
                                   "(data_type || decode(data_type, 'CHAR', '(', 'NUMBER', '(', 'VARCHAR', '(', 'VARCHAR2', '(', '') || " & _
                                   "decode(data_type, 'NUMBER', to_char(data_precision) || decode(data_scale, null, '', ',' || to_char(data_scale)) " & _
                                   ", 'DATE', '', 'LONG', '', 'LONGRAW','', TO_CHAR(data_length)) ||  " & _
                                   "decode(data_type, 'CHAR', ')', 'NUMBER', ')', 'VARCHAR', ')', 'VARCHAR2', ')', ''))) " & _
                                   "as Type " & _
                                "from user_tab_cols where table_name = " & cmdPa.AP(Trim(tableNameArray(1))) & " " & _
                                "order by COLUMN_ID"

                    dt = gDB.getDataTable(runCmdSql, cnn, transaction, , cmdPa, msgLog)


                    dtdiv.Controls.Clear()

                    resTable.BorderStyle = BorderStyle.None
                    resTable.CellPadding = 1
                    resTable.CellSpacing = 1
                    resTable.HorizontalAlign = HorizontalAlign.Center
                    resTable.Width = Unit.Percentage(95)

                    dtdiv.Controls.Add(resTable)

                    If dt.Rows.Count = 0 Then
                        tmpTR = New TableRow
                        tmpTD = New TableCell
                        tmpTD.Text = "No record found!"
                        tmpTR.Controls.Add(tmpTD)
                        resTable.Controls.Add(tmpTR)
                    Else
                        tmpTR = New TableRow
                        tmpTD = New TableCell
                        tmpTR.Controls.Add(tmpTD)
                        resTable.Controls.Add(tmpTR)

                        tmpGV = genResultGV(dt)

                        tmpGV.DataSource = dt
                        tmpGV.DataBind()

                        tmpTD.Controls.Add(tmpGV)
                    End If

                Case "SYSTEM"
                    uiFun.displayMsg(Me, "", "System command not support!", Session("gLang"))
                    Exit Sub

            End Select



            If isCommit Then
                transaction.Commit()
            Else
                transaction.Rollback()
            End If

        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

            dtdiv.Controls.Clear()

            resTable.BorderStyle = BorderStyle.None
            resTable.CellPadding = 1
            resTable.CellSpacing = 1
            resTable.HorizontalAlign = HorizontalAlign.Center
            resTable.Width = Unit.Percentage(95)

            dtdiv.Controls.Add(resTable)

            tmpTR = New TableRow
            tmpTD = New TableCell
            tmpTD.Text = ex.Message
            tmpTR.Controls.Add(tmpTD)
            resTable.Controls.Add(tmpTR)

        Finally
            'If Not transaction Is Nothing Then
            '    transaction.Rollback()
            '    transaction = Nothing
            'End If
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Function genResultGV(ByRef dt As DataTable) As GridView
        Dim dtGrid As New GridView
        Dim tmpBndField As BoundField
        Dim i As Integer

        For i = 0 To dt.Columns.Count - 1
            tmpBndField = New BoundField
            tmpBndField.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
            tmpBndField.DataField = dt.Columns(i).ColumnName
            tmpBndField.HeaderText = dt.Columns(i).ColumnName
            tmpBndField.ItemStyle.CssClass = "GV"
            tmpBndField.ItemStyle.Wrap = False
            'tmpBndField.ControlStyle.Width = Unit.Percentage(7)
            'tmpBndField.ItemStyle.Width = Unit.Percentage(7)
            dtGrid.Columns.Add(tmpBndField)
        Next


        'Set grid view style
        dtGrid.Font.Name = "Arial"
        dtGrid.Font.Overline = False
        dtGrid.AutoGenerateColumns = False

        'dtGrid.PageSize = gU.decodeEmptyCInt(srch_table.Rows(0).Item("srch_page_size").ToString, "20")

        dtGrid.BorderColor = Drawing.Color.Silver
        dtGrid.BorderStyle = BorderStyle.Solid
        dtGrid.BorderWidth = 1
        dtGrid.CellPadding = 3
        dtGrid.CaptionAlign = TableCaptionAlign.Top

        'dtGrid.RowStyle.BorderColor = Drawing.Color.Silver
        'dtGrid.RowStyle.BorderStyle = BorderStyle.Solid
        'dtGrid.RowStyle.BorderWidth = 1

        'dtGrid.HeaderStyle.BorderColor = Drawing.Color.Silver
        'dtGrid.HeaderStyle.BorderStyle = BorderStyle.Inset
        'dtGrid.HeaderStyle.BorderWidth = 1
        dtGrid.HeaderStyle.Font.Name = "Arial"
        dtGrid.HeaderStyle.Font.Bold = True
        dtGrid.HeaderStyle.Font.Size = 9
        dtGrid.HeaderStyle.HorizontalAlign = HorizontalAlign.Left
        dtGrid.HeaderStyle.VerticalAlign = VerticalAlign.Middle
        dtGrid.HeaderStyle.Wrap = True
        dtGrid.HeaderStyle.CssClass = "DtlLabel"

        dtGrid.FooterStyle.BackColor = Drawing.Color.White
        dtGrid.SelectedRowStyle.Font.Bold = True
        dtGrid.SelectedRowStyle.ForeColor = Drawing.Color.White

        dtGrid.Width = Unit.Percentage(100)

        Return dtGrid
    End Function

End Class
