Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing

Partial Class ConvertImage
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private DDFORMAT As String = ""
    Private PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR")


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then

        End If

    End Sub

    Protected Sub btnConvert_Click(sender As Object, e As System.EventArgs) Handles btnConvert.Click
        Dim selectSQL As String = " SELECT IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ITM_PICTURE1, ITM_PICTURE2 " & _
                                  " FROM WMS_ITEM where ((ITM_PICTURE1 is not null and ITM_PICTURE1 <> '') or  (ITM_PICTURE2 is not null and ITM_PICTURE2 <> '')) and ITM_PHOTO1 is null order by itm_code"
        Dim updateSQL As String = ""
        Dim imageConvertCount As Integer = 0
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim gConn As SqlConnection = gDB.getConnection
        'Dim transaction As SqlTransaction = Nothing
        Dim successFlag As Boolean = False
        Dim dt As DataTable

        Try
            gConn = gDB.getConnection()
            'transaction = gConn.BeginTransaction()

            dt = gDB.getDataTable(selectSQL, gConn)

            If dt.Rows.Count > 0 Then              
                For i = 0 To dt.Rows.Count - 1                                        
                    Dim imp_code, storer_code, itm_code, pack_key As String

                    Dim newjpegByteSize As Long = 0

                    imp_code = dt.Rows(i).Item("IMP_CODE").ToString.Trim
                    storer_code = dt.Rows(i).Item("STORER_CODE").ToString.Trim
                    itm_code = dt.Rows(i).Item("ITM_CODE").ToString.Trim
                    pack_key = dt.Rows(i).Item("PACK_KEY").ToString.Trim

                    If dt.Rows(i).Item("ITM_PICTURE1").ToString.Trim <> "" Then
                        If IO.File.Exists(PHY_PS_DIR & "\MAST_IM\" & dt.Rows(i).Item("ITM_PICTURE1").ToString.Trim) Then
                            Dim imageConverter As New ImageConverter()
                            Dim imageByte As Byte() = Nothing
                            Dim Path As String = ""

                            Path = PHY_PS_DIR & "\MAST_IM\" & dt.Rows(i).Item("ITM_PICTURE1").ToString.Trim

                            Dim lImage As Image = System.Drawing.Image.FromFile(Path)

                            Dim jpegByteSize As Long = (New System.IO.FileInfo(Path)).Length

                            If jpegByteSize > 716800 Then
                                Dim _oGenerator As New ThumbGenerator
                                Dim bitmap As Bitmap

                                _oGenerator.SetParams(Path, lImage.Width, lImage.Height, False, False, True)
                                bitmap = _oGenerator.ExtractThumbnail

                                Dim myImageCodecInfo As ImageCodecInfo
                                Dim myEncoderParameters As System.Drawing.Imaging.EncoderParameters
                                myImageCodecInfo = GetEncoderInfo("image/jpeg")

                                myEncoderParameters = New EncoderParameters(1)
                                myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 80L)

                                Using ms = New MemoryStream(jpegByteSize)
                                    bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                    newjpegByteSize = ms.Length
                                    ms.Close()
                                    ms.Dispose()
                                End Using

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 70L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 60L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 50L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                imageByte = DirectCast(ImageConverter.ConvertTo(bitmap, GetType(Byte())), Byte())

                                bitmap.Dispose()
                            Else
                                imageByte = DirectCast(ImageConverter.ConvertTo(lImage, GetType(Byte())), Byte())
                            End If

                            lImage.Dispose()

                            paP = New GlobalDBFunc.DBCmdPara
                            updateSQL = "update wms_item set ITM_PHOTO1=" & paP.AP(imageByte, SqlDbType.VarBinary) & _
                                        " WHERE IMP_CODE=" & paP.AP(imp_code) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(storer_code)) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(itm_code)) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(pack_key))
                            gDB.amendData(updateSQL, gConn, Nothing, paP)

                            imageConvertCount += 1
                        End If
                    End If

                    newjpegByteSize = 0

                    If dt.Rows(i).Item("ITM_PICTURE2").ToString.Trim <> "" Then
                        If IO.File.Exists(PHY_PS_DIR & "\MAST_IM\" & dt.Rows(i).Item("ITM_PICTURE2").ToString.Trim) Then
                            Dim imageConverter As New ImageConverter()
                            Dim imageByte As Byte() = Nothing
                            Dim Path As String = ""

                            Path = PHY_PS_DIR & "\MAST_IM\" & dt.Rows(i).Item("ITM_PICTURE2").ToString.Trim

                            Dim lImage As Image = System.Drawing.Image.FromFile(Path)

                            Dim jpegByteSize As Long = (New System.IO.FileInfo(Path)).Length

                            If jpegByteSize > 716800 Then
                                Dim _oGenerator As New ThumbGenerator
                                Dim bitmap As Bitmap

                                _oGenerator.SetParams(Path, lImage.Width, lImage.Height, False, False, True)
                                bitmap = _oGenerator.ExtractThumbnail

                                Dim myImageCodecInfo As ImageCodecInfo
                                Dim myEncoderParameters As System.Drawing.Imaging.EncoderParameters
                                myImageCodecInfo = GetEncoderInfo("image/jpeg")

                                myEncoderParameters = New EncoderParameters(1)
                                myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 80L)

                                Using ms = New MemoryStream(jpegByteSize)
                                    bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                    newjpegByteSize = ms.Length
                                    ms.Dispose()
                                End Using

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 70L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 60L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                If newjpegByteSize > 716800 Then
                                    myEncoderParameters = New EncoderParameters(1)
                                    myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 50L)

                                    Using ms = New MemoryStream(jpegByteSize)
                                        bitmap.Save(ms, myImageCodecInfo, myEncoderParameters)
                                        newjpegByteSize = ms.Length
                                        ms.Close()
                                        ms.Dispose()
                                    End Using
                                End If

                                imageByte = DirectCast(ImageConverter.ConvertTo(bitmap, GetType(Byte())), Byte())

                                bitmap.Dispose()
                            Else
                                imageByte = DirectCast(ImageConverter.ConvertTo(lImage, GetType(Byte())), Byte())
                            End If

                            lImage.Dispose()

                            paP = New GlobalDBFunc.DBCmdPara
                            updateSQL = "update wms_item set ITM_PHOTO2=" & paP.AP(imageByte, SqlDbType.VarBinary) & _
                                        " WHERE IMP_CODE=" & paP.AP(imp_code) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(storer_code)) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(itm_code)) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(pack_key))
                            gDB.amendData(updateSQL, gConn, Nothing, paP)

                            imageConvertCount += 1
                        End If
                    End If

                    GC.Collect()
                Next
            End If

            'gDB.amendData(updateSQL, gConn, transaction, paP)

            'transaction.Commit()
            successFlag = True

        Catch ex As Exception
            ' transaction.Rollback()
            ' transaction = Nothing
            Response.Write(ex.Message)
        
        Finally

            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

        If successFlag Then
            resultlbl.Text = imageConvertCount & " images Converted."
        End If

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
