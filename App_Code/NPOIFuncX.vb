Imports Microsoft.VisualBasic
Imports NPOI.XSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.SS.UserModel

Public Class NPOIFuncX

    Public Sub CopyCell(ByRef oldCell As XSSFCell, ByRef newCell As XSSFCell, Optional ByVal styleOnly As Boolean = False)
        'If the old cell is null, exit
        If oldCell Is Nothing Then
            Exit Sub
        End If

        Dim workbook As XSSFWorkbook = newCell.Sheet.Workbook

        'Copy style from old cell and apply to new cell
        'Dim newCellStyle As XSSFCellStyle = workbook.CreateCellStyle()
        'newCellStyle.CloneStyleFrom(oldCell.CellStyle)
        'newCell.CellStyle = newCellStyle

        newCell.CellStyle = oldCell.CellStyle

        'If there is a cell comment, copy
        If oldCell.CellComment IsNot Nothing Then
            newCell.CellComment = oldCell.CellComment
        End If

        'Set the cell data type
        newCell.SetCellType(oldCell.CellType)

        If Not styleOnly Then
            'If there is a cell hyperlink, copy
            If oldCell.Hyperlink IsNot Nothing Then
                newCell.Hyperlink = oldCell.Hyperlink
            End If

            'Set the cell data value
            Select Case (oldCell.CellType)
                Case CellType.BLANK
                    'newCell.SetCellValue(oldCell.StringCellValue)
                    Exit Select

                Case CellType.BOOLEAN
                    newCell.SetCellValue(oldCell.BooleanCellValue)
                    Exit Select

                Case CellType.ERROR
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue)
                    Exit Select

                Case CellType.FORMULA
                    newCell.CellFormula = oldCell.CellFormula
                    Exit Select

                Case CellType.NUMERIC
                    newCell.SetCellValue(oldCell.NumericCellValue)
                    Exit Select

                Case CellType.STRING
                    newCell.SetCellValue(oldCell.RichStringCellValue)
                    Exit Select

                Case CellType.Unknown
                    newCell.SetCellValue(oldCell.StringCellValue)
                    Exit Select
            End Select
        End If


    End Sub

    Public Sub CopyRange(ByRef sourceSheet As XSSFSheet, ByRef destSheet As XSSFSheet, ByRef oldRange As CellRangeAddress, ByVal destRow As Integer, ByVal destCol As Integer, Optional ByVal styleOnly As Boolean = False)
        Dim rowDiff As Integer = oldRange.FirstRow - destRow
        Dim colDiff As Integer = oldRange.FirstColumn - destCol
        'Dim newRange As New CellRangeAddress(destRow, oldRange.LastRow - rowDiff, destCol, oldRange.LastColumn - colDiff)

        For i As Integer = oldRange.FirstRow To oldRange.LastRow
            For j As Integer = oldRange.FirstColumn To oldRange.LastColumn
                Dim sourceRow As XSSFRow = sourceSheet.GetRow(i)
                If sourceRow IsNot Nothing Then
                    Dim oldCell As XSSFCell = sourceSheet.GetRow(i).GetCell(j)
                    If oldCell IsNot Nothing Then
                        Dim newRow As XSSFRow = destSheet.GetRow(i - rowDiff)
                        If newRow Is Nothing Then
                            newRow = destSheet.CreateRow(i - rowDiff)
                        End If
                        Dim newCell As XSSFCell = newRow.GetCell(j - colDiff)
                        If newCell Is Nothing Then
                            newCell = newRow.CreateCell(j - colDiff)
                        End If
                        CopyCell(oldCell, newCell, styleOnly)
                    End If
                End If
            Next
        Next

        'If there are are any merged regions in the source row, copy to new row
        For i As Integer = 0 To sourceSheet.NumMergedRegions - 1
            Dim cellRangeAddress As CellRangeAddress = sourceSheet.GetMergedRegion(i)

            If cellRangeAddress.FirstRow >= oldRange.FirstRow AndAlso cellRangeAddress.FirstRow <= oldRange.LastRow AndAlso _
                cellRangeAddress.FirstColumn >= oldRange.FirstColumn AndAlso cellRangeAddress.FirstColumn <= oldRange.LastColumn Then

                Dim newCellRangeAddress As New CellRangeAddress(cellRangeAddress.FirstRow - rowDiff, cellRangeAddress.LastRow - rowDiff, cellRangeAddress.FirstColumn - colDiff, cellRangeAddress.LastColumn - colDiff)

                destSheet.AddMergedRegion(newCellRangeAddress)
            End If
        Next
    End Sub

    Public Sub CopyRow(ByRef workbook As XSSFWorkbook, ByRef sourceSheet As XSSFSheet, ByRef destSheet As XSSFSheet, ByVal sourceRowNum As Integer, ByVal destRowNum As Integer, Optional ByVal styleOnly As Boolean = False)
        'Get the source / new row
        Dim newRow As XSSFRow = destSheet.GetRow(destRowNum)
        Dim sourceRow As XSSFRow = sourceSheet.GetRow(sourceRowNum)

        'If the row exist in destination, push down all rows by 1 else create a new row
        If newRow IsNot Nothing Then
            destSheet.ShiftRows(destRowNum, destSheet.LastRowNum, 1)
        Else
            newRow = destSheet.CreateRow(destRowNum)
        End If

        'Loop through source columns to add to new row
        For i As Integer = 0 To sourceRow.LastCellNum - 1
            'Grab a copy of the old/new cell
            Dim oldCell As XSSFCell = sourceRow.GetCell(i)
            If oldCell IsNot Nothing Then
                Dim newCell As XSSFCell = newRow.CreateCell(i)

                CopyCell(oldCell, newCell, styleOnly)
            End If
        Next

        'If there are are any merged regions in the source row, copy to new row
        For i As Integer = 0 To sourceSheet.NumMergedRegions - 1
            Dim cellRangeAddress As CellRangeAddress = sourceSheet.GetMergedRegion(i)

            If cellRangeAddress.FirstRow = sourceRow.RowNum Then
                Dim newCellRangeAddress As New CellRangeAddress(newRow.RowNum, (newRow.RowNum + (cellRangeAddress.FirstRow - cellRangeAddress.LastRow)), cellRangeAddress.FirstColumn, cellRangeAddress.LastColumn)

                destSheet.AddMergedRegion(newCellRangeAddress)
            End If
        Next

    End Sub

    Public Function getCell(ByRef ExcelWSObj As XSSFSheet, ByVal rowNo As Integer, ByVal colNo As Integer) As XSSFCell
        Dim ExcelRow As XSSFRow
        Dim ExcelCell As XSSFCell

        ExcelRow = ExcelWSObj.GetRow(rowNo)
        If ExcelRow Is Nothing Then
            ExcelRow = ExcelWSObj.CreateRow(rowNo)
        End If

        ExcelCell = ExcelRow.GetCell(colNo)
        If ExcelCell Is Nothing Then
            ExcelCell = ExcelRow.CreateCell(colNo)
        End If

        Return ExcelCell
    End Function

    Public Sub setCellValue(ByRef ExcelWSObj As XSSFSheet, ByVal rowNo As Integer, ByVal colNo As Integer, ByRef cellValue As String, Optional ByVal itemType As String = "")
        Dim ExcelRow As XSSFRow
        Dim ExcelCell As XSSFCell

        'If isCreate Then
        '    ExcelRow = ExcelWSObj.CreateRow(rowNo)
        'Else
        '    ExcelRow = ExcelWSObj.GetRow(rowNo)
        'End If

        If cellValue = "" Then
            Exit Sub
        End If

        ExcelRow = ExcelWSObj.GetRow(rowNo)
        If ExcelRow Is Nothing Then
            ExcelRow = ExcelWSObj.CreateRow(rowNo)
        End If

        ExcelCell = ExcelRow.GetCell(colNo)
        If ExcelCell Is Nothing Then
            ExcelCell = ExcelRow.CreateCell(colNo)
        End If

        If itemType = "INT" Then
            ExcelCell.SetCellValue(CLng(cellValue))
        ElseIf itemType = "DEC" Then
            'ExcelCell.SetCellValue(CDbl(drItem))
            ExcelCell.SetCellValue(Val(cellValue))
        ElseIf itemType = "CHKLIST" Then
            Dim tmpValue As String

            tmpValue = cellValue

            If tmpValue <> "" Then
                tmpValue = Replace(tmpValue, "/><", ",")
                tmpValue = Mid(tmpValue, 2, Len(tmpValue) - 3)

                ExcelCell.SetCellValue(tmpValue)
            End If
        ElseIf itemType = "FORMULA" Then
            ExcelCell.CellFormula = cellValue
        Else
            ExcelCell.SetCellValue(cellValue)
        End If
    End Sub

    Public Function getCellName(ByVal rowNo As Integer, ByVal colNo As Integer) As String
        Return getXLCol(colNo) & (rowNo + 1)
    End Function

    Public Function getXLCol(ByVal Col As Integer) As String
        ' Col is the present column, not the number of cols
        Const A As Integer = 65    'ASCII value for capital A
        Dim sCol As String
        Dim iRemain As Integer
        ' THIS ALGORITHM ONLY WORKS UP TO ZZ. It fails on AAA
        If Col > 701 Then
            getXLCol = ""
            Exit Function
        End If
        If Col <= 25 Then
            sCol = Chr(A + Col)
        Else
            iRemain = Int((Col / 26)) - 1
            sCol = Chr(A + iRemain) & getXLCol(Col Mod 26)
        End If
        getXLCol = sCol
    End Function


    Public Sub InsertRows(ByRef sheet1 As XSSFSheet, fromRowIndex As Integer, rowCount As Integer)
        sheet1.ShiftRows(fromRowIndex, sheet1.LastRowNum, rowCount, True, False)

        For rowIndex As Integer = fromRowIndex To fromRowIndex + (rowCount - 1)
            Dim rowSource As XSSFRow = sheet1.GetRow(rowIndex + rowCount)
            Dim rowInsert As XSSFRow = sheet1.CreateRow(rowIndex)
            'rowInsert.Height = rowSource.Height
            For colIndex As Integer = 0 To rowSource.LastCellNum - 1
                Dim cellSource As XSSFCell = rowSource.GetCell(colIndex)
                Dim cellInsert As XSSFCell = rowInsert.CreateCell(colIndex)
                If cellSource IsNot Nothing Then
                    cellInsert.CellStyle = cellSource.CellStyle
                End If
            Next
        Next
    End Sub

    Public Sub removeRow(sheet As XSSFSheet, rowIndex As Integer)
        Dim lastRowNum As Integer = sheet.LastRowNum
        If rowIndex >= 0 AndAlso rowIndex < lastRowNum Then
            sheet.ShiftRows(rowIndex + 1, lastRowNum, -1)
        End If
        If rowIndex = lastRowNum Then
            Dim removingRow As XSSFRow = sheet.GetRow(rowIndex)
            If removingRow IsNot Nothing Then
                sheet.RemoveRow(removingRow)
            End If
        End If
    End Sub
End Class
