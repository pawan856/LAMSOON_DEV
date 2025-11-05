Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web.Caching
Imports System.Configuration
Imports System.Data

Partial Class test
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim myimg As Bitmap = Code128Rendering.MakeBarcodeImage("0000032 1", 2, True)

        Dim myImageCodecInfo As ImageCodecInfo
        myImageCodecInfo = GetEncoderInfo("image/jpeg")

        Dim myEncoderParameters As System.Drawing.Imaging.EncoderParameters
        myEncoderParameters = New EncoderParameters(1)
        myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 100L)

        myimg.Save(Context.Response.OutputStream, myImageCodecInfo, myEncoderParameters)

        Context.Response.OutputStream.Close()
        Context.Response.End()
        myimg.Dispose()
    End Sub

    Private Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim j As Integer
        Dim encoders() As ImageCodecInfo
        encoders = ImageCodecInfo.GetImageEncoders()

        j = 0
        While j < encoders.Length
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
            j += 1
        End While
        Return Nothing

    End Function
End Class
