<%@ WebHandler Language="VB" Class="FileHandler" %>

Imports System
Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Public Class FileHandler : Implements IHttpHandler
    Private _path As String
    Private _filename As String
    Private TEMP_FILEPATH_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        _path = System.Web.HttpUtility.UrlDecode(context.Request("filePath").ToString)
        _filename = System.Web.HttpUtility.UrlDecode(context.Request("fileName").ToString)
        '_path = TEMP_FILEPATH_DIR & "\" & _filename
        
        Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(_path)
        
        If File.Exists(_path) Then
            context.Response.Clear()
            context.Response.Buffer = True
            context.Response.AddHeader("Content-Disposition", "attachment; filename=" & _filename)
            context.Response.AddHeader("Content-Length", nFile.Length.ToString())
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.WriteFile(nFile.FullName)
            context.Response.Flush()
            context.Response.End()
        End If
        
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property

End Class