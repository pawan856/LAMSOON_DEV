Imports Microsoft.VisualBasic

Imports System.Drawing
Imports System.Diagnostics
Imports System.IO
Imports System.Drawing.Imaging
Imports ZXing

Public NotInheritable Class Code128Rendering
    Private Sub New()
    End Sub

    Private Shared ReadOnly cPatterns As Integer(,) = _
        { _
        {2, 1, 2, 2, 2, 2, 0, 0}, _
        {2, 2, 2, 1, 2, 2, 0, 0}, _
        {2, 2, 2, 2, 2, 1, 0, 0}, _
        {1, 2, 1, 2, 2, 3, 0, 0}, _
        {1, 2, 1, 3, 2, 2, 0, 0}, _
        {1, 3, 1, 2, 2, 2, 0, 0}, _
        {1, 2, 2, 2, 1, 3, 0, 0}, _
        {1, 2, 2, 3, 1, 2, 0, 0}, _
        {1, 3, 2, 2, 1, 2, 0, 0}, _
        {2, 2, 1, 2, 1, 3, 0, 0}, _
        {2, 2, 1, 3, 1, 2, 0, 0}, _
        {2, 3, 1, 2, 1, 2, 0, 0}, _
        {1, 1, 2, 2, 3, 2, 0, 0}, _
        {1, 2, 2, 1, 3, 2, 0, 0}, _
        {1, 2, 2, 2, 3, 1, 0, 0}, _
        {1, 1, 3, 2, 2, 2, 0, 0}, _
        {1, 2, 3, 1, 2, 2, 0, 0}, _
        {1, 2, 3, 2, 2, 1, 0, 0}, _
        {2, 2, 3, 2, 1, 1, 0, 0}, _
        {2, 2, 1, 1, 3, 2, 0, 0}, _
        {2, 2, 1, 2, 3, 1, 0, 0}, _
        {2, 1, 3, 2, 1, 2, 0, 0}, _
        {2, 2, 3, 1, 1, 2, 0, 0}, _
        {3, 1, 2, 1, 3, 1, 0, 0}, _
        {3, 1, 1, 2, 2, 2, 0, 0}, _
        {3, 2, 1, 1, 2, 2, 0, 0}, _
        {3, 2, 1, 2, 2, 1, 0, 0}, _
        {3, 1, 2, 2, 1, 2, 0, 0}, _
        {3, 2, 2, 1, 1, 2, 0, 0}, _
        {3, 2, 2, 2, 1, 1, 0, 0}, _
        {2, 1, 2, 1, 2, 3, 0, 0}, _
        {2, 1, 2, 3, 2, 1, 0, 0}, _
        {2, 3, 2, 1, 2, 1, 0, 0}, _
        {1, 1, 1, 3, 2, 3, 0, 0}, _
        {1, 3, 1, 1, 2, 3, 0, 0}, _
        {1, 3, 1, 3, 2, 1, 0, 0}, _
        {1, 1, 2, 3, 1, 3, 0, 0}, _
        {1, 3, 2, 1, 1, 3, 0, 0}, _
        {1, 3, 2, 3, 1, 1, 0, 0}, _
        {2, 1, 1, 3, 1, 3, 0, 0}, _
        {2, 3, 1, 1, 1, 3, 0, 0}, _
        {2, 3, 1, 3, 1, 1, 0, 0}, _
        {1, 1, 2, 1, 3, 3, 0, 0}, _
        {1, 1, 2, 3, 3, 1, 0, 0}, _
        {1, 3, 2, 1, 3, 1, 0, 0}, _
        {1, 1, 3, 1, 2, 3, 0, 0}, _
        {1, 1, 3, 3, 2, 1, 0, 0}, _
        {1, 3, 3, 1, 2, 1, 0, 0}, _
        {3, 1, 3, 1, 2, 1, 0, 0}, _
        {2, 1, 1, 3, 3, 1, 0, 0}, _
        {2, 3, 1, 1, 3, 1, 0, 0}, _
        {2, 1, 3, 1, 1, 3, 0, 0}, _
        {2, 1, 3, 3, 1, 1, 0, 0}, _
        {2, 1, 3, 1, 3, 1, 0, 0}, _
        {3, 1, 1, 1, 2, 3, 0, 0}, _
        {3, 1, 1, 3, 2, 1, 0, 0}, _
        {3, 3, 1, 1, 2, 1, 0, 0}, _
        {3, 1, 2, 1, 1, 3, 0, 0}, _
        {3, 1, 2, 3, 1, 1, 0, 0}, _
        {3, 3, 2, 1, 1, 1, 0, 0}, _
        {3, 1, 4, 1, 1, 1, 0, 0}, _
        {2, 2, 1, 4, 1, 1, 0, 0}, _
        {4, 3, 1, 1, 1, 1, 0, 0}, _
        {1, 1, 1, 2, 2, 4, 0, 0}, _
        {1, 1, 1, 4, 2, 2, 0, 0}, _
        {1, 2, 1, 1, 2, 4, 0, 0}, _
        {1, 2, 1, 4, 2, 1, 0, 0}, _
        {1, 4, 1, 1, 2, 2, 0, 0}, _
        {1, 4, 1, 2, 2, 1, 0, 0}, _
        {1, 1, 2, 2, 1, 4, 0, 0}, _
        {1, 1, 2, 4, 1, 2, 0, 0}, _
        {1, 2, 2, 1, 1, 4, 0, 0}, _
        {1, 2, 2, 4, 1, 1, 0, 0}, _
        {1, 4, 2, 1, 1, 2, 0, 0}, _
        {1, 4, 2, 2, 1, 1, 0, 0}, _
        {2, 4, 1, 2, 1, 1, 0, 0}, _
        {2, 2, 1, 1, 1, 4, 0, 0}, _
        {4, 1, 3, 1, 1, 1, 0, 0}, _
        {2, 4, 1, 1, 1, 2, 0, 0}, _
        {1, 3, 4, 1, 1, 1, 0, 0}, _
        {1, 1, 1, 2, 4, 2, 0, 0}, _
        {1, 2, 1, 1, 4, 2, 0, 0}, _
        {1, 2, 1, 2, 4, 1, 0, 0}, _
        {1, 1, 4, 2, 1, 2, 0, 0}, _
        {1, 2, 4, 1, 1, 2, 0, 0}, _
        {1, 2, 4, 2, 1, 1, 0, 0}, _
        {4, 1, 1, 2, 1, 2, 0, 0}, _
        {4, 2, 1, 1, 1, 2, 0, 0}, _
        {4, 2, 1, 2, 1, 1, 0, 0}, _
        {2, 1, 2, 1, 4, 1, 0, 0}, _
        {2, 1, 4, 1, 2, 1, 0, 0}, _
        {4, 1, 2, 1, 2, 1, 0, 0}, _
        {1, 1, 1, 1, 4, 3, 0, 0}, _
        {1, 1, 1, 3, 4, 1, 0, 0}, _
        {1, 3, 1, 1, 4, 1, 0, 0}, _
        {1, 1, 4, 1, 1, 3, 0, 0}, _
        {1, 1, 4, 3, 1, 1, 0, 0}, _
        {4, 1, 1, 1, 1, 3, 0, 0}, _
        {4, 1, 1, 3, 1, 1, 0, 0}, _
        {1, 1, 3, 1, 4, 1, 0, 0}, _
        {1, 1, 4, 1, 3, 1, 0, 0}, _
        {3, 1, 1, 1, 4, 1, 0, 0}, _
        {4, 1, 1, 1, 3, 1, 0, 0}, _
        {2, 1, 1, 4, 1, 2, 0, 0}, _
        {2, 1, 1, 2, 1, 4, 0, 0}, _
        {2, 1, 1, 2, 3, 2, 0, 0}, _
        {2, 3, 3, 1, 1, 1, 2, 0} _
        }

    Private Const cQuietWidth As Integer = 10

    Public Shared Function MakeBarcodeImage(InputData As String, BarWeight As Integer, AddQuietZone As Boolean) As Image
        ' get the Code128 codes to represent the message
        Dim content As New Code128Content(InputData)
        Dim codes As Integer() = content.Codes

        Dim width As Integer, height As Integer
        width = ((codes.Length - 3) * 11 + 35) * BarWeight
        height = Convert.ToInt32(System.Math.Ceiling(Convert.ToSingle(width) * 0.15F))

        If AddQuietZone Then
            ' on both sides
            width += 2 * cQuietWidth * BarWeight
        End If

        ' get surface to draw on
        Dim myimg As Image = New System.Drawing.Bitmap(width, height)
        Using gr As Graphics = Graphics.FromImage(myimg)

            ' set to white so we don't have to fill the spaces with white
            gr.FillRectangle(System.Drawing.Brushes.White, 0, 0, width, height)

            ' skip quiet zone
            Dim cursor As Integer = If(AddQuietZone, cQuietWidth * BarWeight, 0)

            For codeidx As Integer = 0 To codes.Length - 1
                Dim code As Integer = codes(codeidx)

                ' take the bars two at a time: a black and a white
                For bar As Integer = 0 To 7 Step 2
                    Dim barwidth As Integer = cPatterns(code, bar) * BarWeight
                    Dim spcwidth As Integer = cPatterns(code, bar + 1) * BarWeight

                    ' if width is zero, don't try to draw it
                    If barwidth > 0 Then
                        gr.FillRectangle(System.Drawing.Brushes.Black, cursor, 0, barwidth, height)
                    End If

                    ' note that we never need to draw the space, since we 
                    ' initialized the graphics to all white

                    ' advance cursor beyond this pair
                    cursor += (barwidth + spcwidth)
                Next
            Next
        End Using

        Return myimg

    End Function

    Public Shared Function MakeQRImage(InputData As String, ImagePath As String) As Image
        Dim writer = New BarcodeWriter()
        writer.Format = BarcodeFormat.QR_CODE
        Dim result = writer.Write(InputData)
        Dim path As String = ImagePath
        Dim barcodeBitmap = New Bitmap(result)

        Using memory As New MemoryStream()
            Using fs As New FileStream(path, FileMode.Create, FileAccess.ReadWrite)
                barcodeBitmap.Save(memory, ImageFormat.Jpeg)
                Dim bytes As Byte() = memory.ToArray()
                fs.Write(bytes, 0, bytes.Length)
            End Using
        End Using
        Return barcodeBitmap
    End Function
End Class