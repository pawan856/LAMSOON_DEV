<%@ WebHandler Language="VB" Class="ThumbnailHandler" %>

Imports System
Imports System.IO
Imports System.Web
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web.Caching
Imports System.Configuration
Imports ThumbGenerator
Imports System.Data


Public Class ThumbnailHandler : Implements IHttpHandler, System.Web.SessionState.IRequiresSessionState  
    
    Private _path As String
    Private _width, _height As Integer
    Private _bStretch, _bBevel, _orgnsize, _savethumb As Boolean
    Private _oGenerator As ThumbGenerator
    Private _vm As String = ""
    
    Private ServerIP As String = System.Configuration.ConfigurationManager.AppSettings.Item("WEB_SERVER_IP")
    Private PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR")
    Private SYSP_UPLD_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_UPLD_DIR")
    Private PHY_THUMBNAIL_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_THUMBNAIL_DIR")
    
    Private ar As New AccessRightUtils
    
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim bitmap As Bitmap
        Dim sCacheKey As String
        Dim bFoundInCache As Boolean = True
        
        If context.Session("usr_id") Is Nothing Or context.Session("usr_id") = "" Then
                context.Session.Remove("PAGE_SESSION_MENU_CODE")
                Exit Sub
            End If
        
        If Not context.Request("code") Is Nothing Then
            PHY_PS_DIR = PHY_PS_DIR & "\" & context.Request("code")
            SYSP_UPLD_DIR = SYSP_UPLD_DIR
            PHY_THUMBNAIL_DIR = PHY_THUMBNAIL_DIR & "\" & context.Request("code")
        End if      

        _oGenerator = New ThumbGenerator

        If Not context.Request("Width") Is Nothing Then
            _width = Int32.Parse(context.Request("Width"))
        Else
            Try
                _width = Integer.Parse("150")
            Catch ex As ArgumentNullException
                _width = 150
            End Try
        End If
        
        If Not context.Request("Height") Is Nothing Then
            _height = Int32.Parse(context.Request("Height"))
        Else
            Try
                _height = Integer.Parse("150")
            Catch ex As ArgumentNullException
                _height = 150
            End Try
        End If
        
        If Not context.Request("vm") Is Nothing Then
            _vm = context.Request("vm").ToString
        End If
        
        If Not context.Request("savethumb") Is Nothing Then
            _savethumb = CType(context.Request("savethumb"), Boolean)
        Else
            _savethumb = False
        End If

        Const NoThumbFile As String = "no_picture.gif"
        Dim sNoThumbPath As String = context.Request.MapPath("~/images") & "\" & NoThumbFile
        
        If Not IO.Directory.Exists(SYSP_UPLD_DIR) Then IO.Directory.CreateDirectory(SYSP_UPLD_DIR)
        If Not IO.Directory.Exists(PHY_PS_DIR) Then IO.Directory.CreateDirectory(PHY_PS_DIR)
        If Not IO.Directory.Exists(PHY_THUMBNAIL_DIR) Then IO.Directory.CreateDirectory(PHY_THUMBNAIL_DIR)
        
        If Not context.Request("VFilePath") Is Nothing Then
            If _savethumb Then
                _path = PHY_THUMBNAIL_DIR & "\" & context.Request("VFilePath")
            Else
                If Not _vm Is Nothing Then
                    If _vm = "T" Then
                        If Not IO.Directory.Exists(SYSP_UPLD_DIR & "\" & context.Request("userid")) Then IO.Directory.CreateDirectory(SYSP_UPLD_DIR & "\" & context.Request("userid"))
                        _path = SYSP_UPLD_DIR & "\" & context.Request("userid") & "\" & context.Request("VFilePath")
                    Else
                        _path = PHY_PS_DIR & "\" & context.Request("VFilePath")
                    End If
                Else
                    _path = PHY_PS_DIR & "\" & context.Request("VFilePath")
                End If
            End If
        Else
            _path = sNoThumbPath
        End If
        
        If Not File.Exists(_path) Then
            _path = sNoThumbPath
        End If

        _bStretch = context.Request("AllowStretch") = "true"

        _bBevel = context.Request("Bevel") = "true"

        _orgnsize = context.Request("ds") = "true"

        _oGenerator.SetParams(_path, _width, _height, _bStretch, _bBevel, _orgnsize, _vm)

        Dim MyCache As Cache = context.Cache
        sCacheKey = _oGenerator.GetUniqueThumbName

        Dim bRefresh As Boolean = (context.Request("Refresh") = "true")
        
        If bRefresh Then
            MyCache.Remove(sCacheKey)
        End If
        
        If MyCache(sCacheKey) Is Nothing Then
            Try
                bitmap = _oGenerator.ExtractThumbnail
                bFoundInCache = False
            Catch e As Exception
                _path = sNoThumbPath
                _oGenerator.SetParams(_path, _width, _height, _bStretch, False, _orgnsize, _vm)

                sCacheKey = _oGenerator.GetUniqueThumbName
                If MyCache(sCacheKey) Is Nothing Then
                    Try
                        bitmap = _oGenerator.ExtractThumbnail
                        bFoundInCache = False
                    Catch e2 As Exception
                        Return
                    End Try
                Else
                    bitmap = CType(MyCache(sCacheKey), Bitmap)
                End If
            End Try
        Else
            bitmap = CType(MyCache(sCacheKey), Bitmap)
        End If
        
        context.Response.ContentType = "image/jpeg"
        
        Dim myImageCodecInfo As ImageCodecInfo
        myImageCodecInfo = GetEncoderInfo("image/jpeg")
        
        Dim myEncoderParameters As System.Drawing.Imaging.EncoderParameters
        myEncoderParameters = New EncoderParameters(1)
        myEncoderParameters.Param(0) = New EncoderParameter(Encoder.Quality, 60L)
        
        'bitmap.Save(context.Response.OutputStream, ImageFormat.Jpeg)
        
        bitmap.Save(context.Response.OutputStream, myImageCodecInfo, myEncoderParameters)

        Dim bUseCache As Boolean = True
        If (Not bFoundInCache _
                    AndAlso bUseCache) Then
            Dim dependency As CacheDependency = New CacheDependency(_path)
            Dim mins As Integer
            Try
                mins = Integer.Parse("20")
            Catch ex As ArgumentNullException
                mins = 20
            End Try
            
            MyCache.Insert(sCacheKey, bitmap, dependency, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(mins), CacheItemPriority.Default, New CacheItemRemovedCallback(AddressOf RemovedCallback))
            dependency.Dispose()
        End If
        
        context.Response.OutputStream.Close()
        context.Response.End()
        bitmap.Dispose()
    End Sub
    
    Public Shared Sub RemovedCallback(ByVal k As String, ByVal item As Object, ByVal r As CacheItemRemovedReason)
        CType(item, Bitmap).Dispose()
        'LogMessage("Callback")
    End Sub

    Private Shared Sub LogMessage(ByVal mess As String)
        'Dim sw As StreamWriter = New StreamWriter("D:\obi_rpt_server\ASP.NET_log.txt", True)
        'sw.WriteLine(mess)
        'sw.Close()
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