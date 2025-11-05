Imports System.IO
Imports System.Text

Public Class PrgmLog
    Private logFilePath, logFileName, logPrgmLoc As String
    Private logLevel As Integer
    Private logEncoding As System.Text.Encoding
    Public skipLog As Boolean = False

    Private Const LOG_GENERAL As Integer = 0
    Private Const LOG_ERROR As Integer = 1
    Private Const LOG_WARNING As Integer = 2
    Private Const LOG_PROCESS As Integer = 3
    Private Const LOG_XML As Integer = 4
    Private Const LOG_DB As Integer = 5
    Private Const LOG_TIME As Integer = 8
    Private Const LOG_DEBUG As Integer = 9

    Public Sub New()
        logFilePath = ""
        logFileName = ""
        logPrgmLoc = ""
        logLevel = LOG_DB
        logEncoding = Encoding.Unicode
    End Sub

    Public Sub New(ByVal filePath As String, ByVal fileName As String)
        logFilePath = filePath
        logFileName = fileName
        logPrgmLoc = ""
        logLevel = LOG_DB
        logEncoding = Encoding.Unicode
    End Sub

    Public Sub New(ByVal filePath As String, ByVal fileName As String, ByVal defLevel As Integer)
        logFilePath = filePath
        logFileName = fileName
        logPrgmLoc = ""
        logLevel = defLevel
        logEncoding = Encoding.Unicode
    End Sub

    Public Sub New(ByVal filePath As String, ByVal fileName As String, ByVal defLevel As Integer, ByRef encoding As System.Text.Encoding)
        logFilePath = filePath
        logFileName = fileName
        logPrgmLoc = ""
        logLevel = defLevel
        logEncoding = encoding
    End Sub

    Public Property FilePath() As String
        Get
            Return logFilePath
        End Get

        Set(ByVal value As String)
            logFilePath = value
        End Set
    End Property

    Public Property FileName() As String
        Get
            Return logFileName
        End Get

        Set(ByVal value As String)
            logFileName = value
        End Set
    End Property

    Public ReadOnly Property FullFileName() As String
        Get
            Return logFilePath & "\" & logFileName
        End Get
    End Property

    Public Property PrgmLocation() As String
        Get
            Return logPrgmLoc
        End Get

        Set(ByVal value As String)
            logPrgmLoc = value
        End Set
    End Property

    Public Property Encoding() As System.Text.Encoding
        Get
            Return logEncoding
        End Get

        Set(ByVal value As System.Text.Encoding)
            logEncoding = value
        End Set
    End Property

    Public Property DebugLevel() As Integer
        Get
            Return LogLevel
        End Get

        Set(ByVal value As Integer)
            logLevel = value
        End Set
    End Property

    Public Sub writeLog(ByVal message As String)
        If logPrgmLoc <> "" Then
            message = "[" & logPrgmLoc & "] " & message
        End If
        writeSimpleLog("[" & Format(Now, "yyyy-MM-dd hh:mm:ss.fff") & "] " & message)
    End Sub

    Public Sub writeLog(ByVal level As Integer, ByVal message As String)
        If level <= logLevel Then
            writeLog(message)
        End If
    End Sub

    Public Sub writeSimpleLog(ByVal message As String)
        If Not skipLog Then
            Dim logStreamWriter As StreamWriter

            If System.IO.Directory.Exists(logFilePath) Then
                logStreamWriter = New StreamWriter(logFilePath & "\" & logFileName, True, logEncoding)
                logStreamWriter.WriteLine(message)
                logStreamWriter.Close()
            Else
                Throw New System.ApplicationException("Log directory not exist!")
            End If
        End If
    End Sub

    Public Sub writeSimpleLog(ByVal level As Integer, ByVal message As String)
        If level <= logLevel Then
            writeSimpleLog(message)
        End If
    End Sub
End Class
