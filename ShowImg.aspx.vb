Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports System.Drawing.Printing

Partial Class ShowImg
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim imgByte As Byte()

            imgByte = Session("ShowImgByte")
            Dim base64String As String = Convert.ToBase64String(imgByte, 0, imgByte.Length)
            mainImage.ImageUrl = Convert.ToString("data:image/png;base64,") & base64String
        End If
    End Sub
End Class
