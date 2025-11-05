Imports Microsoft.VisualBasic

Public Class AppConfig
    Implements IDisposable

    Private itemRemoved As Boolean = False
    Private reason As CacheItemRemovedReason

    Private disposedValue As Boolean = False

    Public Sub RemovedCallback(ByVal k As String, ByVal v As Object, ByVal r As CacheItemRemovedReason)
        itemRemoved = True
        reason = r
    End Sub

    Public Sub SetCache()
        Dim value As String
        Dim settingKeys As String()
        Dim onRemove As CacheItemRemovedCallback
        Dim PageCache As Cache = HttpContext.Current.Cache

        onRemove = New CacheItemRemovedCallback(AddressOf Me.RemovedCallback)

        settingKeys = System.Configuration.ConfigurationManager.AppSettings.AllKeys

        For x As Integer = 0 To settingKeys.Length - 1
            If PageCache(settingKeys(x)) Is Nothing Then
                value = System.Configuration.ConfigurationManager.AppSettings.Item(settingKeys(x))
                PageCache.Insert(settingKeys(x), value)
            End If
        Next

    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

            End If
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Public Sub RemoveTempSession(ByVal key As String)
        Dim removeKey As New ArrayList

        For i As Integer = 0 To HttpContext.Current.Session.Keys.Count - 1
            If HttpContext.Current.Session.Keys(i).ToString.Contains(key) Then
                removeKey.Add(HttpContext.Current.Session.Keys(i).ToString)
            End If
        Next

        For x As Integer = 0 To removeKey.Count - 1
            HttpContext.Current.Session.Remove(removeKey(x))
        Next
    End Sub
End Class
