Imports Microsoft.VisualBasic
Imports System
Imports System.Text
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Public Class ThumbGenerator
    Private _path As String
    Private _width, _height As Integer
    Private _bStretch, _bBevel, _originalsize As Boolean
    Private _vm As String = ""

    Public Sub SetParams(ByVal path As String, ByVal width As Integer, ByVal height As Integer, _
                            Optional ByVal bStretch As Boolean = False, Optional ByVal bBevel As Boolean = False, _
                            Optional ByVal originalsize As Boolean = False, Optional ByVal vm As String = "")

        _path = path
        _width = width
        _height = height
        _bStretch = bStretch
        _bBevel = bBevel
        _originalsize = originalsize
        _vm = vm

    End Sub

    Public Function ExtractThumbnail() As Bitmap
        Dim bitmapNew As Bitmap
        Dim f As Single
        Dim fx As Single
        Dim fy As Single
        Dim heightTh As Integer
        Dim widthTh As Integer
        Dim heightOrig As Single
        Dim widthOrig As Single

        Dim tempBM As Bitmap

        tempBM = New Bitmap(_path)

        bitmapNew = tempBM

        If _originalsize Then
            widthTh = tempBM.Width
            heightTh = tempBM.Height

            bitmapNew = New Bitmap(tempBM)
            tempBM.Dispose()

            'Return bitmapNew

        Else
            If Not _bStretch Then

                widthOrig = bitmapNew.Width
                heightOrig = bitmapNew.Height
                fx = widthOrig / _width
                fy = heightOrig / _height

                f = Math.Max(fx, fy)
                If f < 1 Then
                    f = 1
                End If
                widthTh = CType(widthOrig / f, Integer)
                heightTh = CType(heightOrig / f, Integer)
            Else
                widthTh = _width
                heightTh = _height
            End If

            bitmapNew = CType(tempBM.GetThumbnailImage(widthTh, heightTh, New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback), IntPtr.Zero), Bitmap)
            tempBM.Dispose()
        End If

        If Not _bBevel Then
            Return bitmapNew
        End If

        Dim heTh As Integer
        Dim widTh As Integer
        widTh = bitmapNew.Width
        heTh = bitmapNew.Height
        Dim Light As Integer = 255
        Dim BevW As Integer = 10
        Dim LowA As Integer = 0
        Dim HighA As Integer = 180
        Dim Dark As Integer = 80

        Dim clrHi1 As Color = Color.FromArgb(LowA, Light, Light, Light)
        Dim clrHi2 As Color = Color.FromArgb(HighA, Light, Light, Light)
        Dim clrDark1 As Color = Color.FromArgb(LowA, Dark, Dark, Dark)
        Dim clrDark2 As Color = Color.FromArgb(HighA, Dark, Dark, Dark)
        Dim br As LinearGradientBrush
        Dim rectSide As Rectangle
        Dim newG As Graphics = Graphics.FromImage(bitmapNew)
        Dim szHorz As Size = New Size(widTh, BevW)
        Dim szVert As Size = New Size(BevW, heTh)

        szHorz = szHorz + New Size(0, 2)
        szVert = szVert + New Size(2, 0)
        rectSide = New Rectangle(New Point(0, (heTh - BevW)), szHorz)
        br = New LinearGradientBrush(rectSide, clrDark1, clrDark2, LinearGradientMode.Vertical)
        rectSide.Inflate(0, -1)
        newG.FillRectangle(br, rectSide)

        rectSide = New Rectangle(New Point((widTh - BevW), 0), szVert)
        br = New LinearGradientBrush(rectSide, clrDark1, clrDark2, LinearGradientMode.Horizontal)
        rectSide.Inflate(-1, 0)
        newG.FillRectangle(br, rectSide)

        szHorz = szHorz - New Size(0, 2)
        szVert = szVert - New Size(2, 0)

        rectSide = New Rectangle(New Point(0, 0), szHorz)
        br = New LinearGradientBrush(rectSide, clrHi2, clrHi1, LinearGradientMode.Vertical)
        newG.FillRectangle(br, rectSide)

        rectSide = New Rectangle(New Point(0, 0), szVert)
        br = New LinearGradientBrush(rectSide, clrHi2, clrHi1, LinearGradientMode.Horizontal)
        newG.FillRectangle(br, rectSide)

        br.Dispose()
        newG.Dispose()

        tempBM.Dispose()
        Return bitmapNew

    End Function

    Public Function GetUniqueThumbName() As String
        Dim sUN As StringBuilder = New StringBuilder

        sUN.AppendFormat("Thumb_{0}_{1}x{2}", Path.GetFileName(_path), _width, _height)

        If _bStretch Then
            sUN.Append("_S")
        End If

        If _bBevel Then
            sUN.Append("_B")
        End If

        If _vm = "T" Then
            sUN.Append("_VM")
        End If

        sUN.Append(".jpg")

        Return sUN.ToString

    End Function

    Public Sub SaveThumbnail(ByVal bmThumb As Bitmap, ByVal path As String)
        bmThumb.Save(path, ImageFormat.Jpeg)
    End Sub

    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function
End Class
