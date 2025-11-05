Imports Microsoft.VisualBasic
Imports System.Text

Public Enum CodeSet
    CodeA
    CodeB
    ' ,CodeC   'not supported
End Enum

Public Class Code128Content
    Private mCodeList As Integer()

    Public Sub New(AsciiData As String)
        mCodeList = StringToCode128(AsciiData)
    End Sub

    Public ReadOnly Property Codes() As Integer()
        Get
            Return mCodeList
        End Get
    End Property

    Private Function StringToCode128(AsciiData As String) As Integer()
        ' turn the string into ascii byte data
        Dim asciiBytes As Byte() = Encoding.ASCII.GetBytes(AsciiData)

        ' decide which codeset to start with
        Dim csa1 As Code128Code.CodeSetAllowed = If(asciiBytes.Length > 0, Code128Code.CodesetAllowedForChar(asciiBytes(0)), Code128Code.CodeSetAllowed.CodeAorB)
        Dim csa2 As Code128Code.CodeSetAllowed = If(asciiBytes.Length > 0, Code128Code.CodesetAllowedForChar(asciiBytes(1)), Code128Code.CodeSetAllowed.CodeAorB)
        Dim currcs As CodeSet = GetBestStartSet(csa1, csa2)

        ' set up the beginning of the barcode
        Dim codes As New System.Collections.ArrayList(asciiBytes.Length + 3)
        ' assume no codeset changes, account for start, checksum, and stop
        codes.Add(Code128Code.StartCodeForCodeSet(currcs))

        ' add the codes for each character in the string
        For i As Integer = 0 To asciiBytes.Length - 1
            Dim thischar As Integer = asciiBytes(i)
            Dim nextchar As Integer = If(asciiBytes.Length > (i + 1), asciiBytes(i + 1), -1)

            codes.AddRange(Code128Code.CodesForChar(thischar, nextchar, currcs))
        Next

        ' calculate the check digit
        Dim checksum As Integer = CInt(codes(0))
        For i As Integer = 1 To codes.Count - 1
            checksum += i * CInt(codes(i))
        Next
        codes.Add(checksum Mod 103)

        codes.Add(Code128Code.StopCode())

        Dim result As Integer() = TryCast(codes.ToArray(GetType(Integer)), Integer())
        Return result
    End Function

    Private Function GetBestStartSet(csa1 As Code128Code.CodeSetAllowed, csa2 As Code128Code.CodeSetAllowed) As CodeSet
        Dim vote As Integer = 0

        vote += If((csa1 = Code128Code.CodeSetAllowed.CodeA), 1, 0)
        vote += If((csa1 = Code128Code.CodeSetAllowed.CodeB), -1, 0)
        vote += If((csa2 = Code128Code.CodeSetAllowed.CodeA), 1, 0)
        vote += If((csa2 = Code128Code.CodeSetAllowed.CodeB), -1, 0)

        Return If((vote > 0), CodeSet.CodeA, CodeSet.CodeB)
        ' ties go to codeB due to my own prejudices
    End Function
End Class

Public NotInheritable Class Code128Code
    Private Sub New()
    End Sub

    Private Const cSHIFT As Integer = 98
    Private Const cCODEA As Integer = 101
    Private Const cCODEB As Integer = 100

    Private Const cSTARTA As Integer = 103
    Private Const cSTARTB As Integer = 104
    Private Const cSTOP As Integer = 106

    Public Shared Function CodesForChar(CharAscii As Integer, LookAheadAscii As Integer, ByRef CurrCodeSet As CodeSet) As Integer()
        Dim result As Integer()
        Dim shifter As Integer = -1

        If Not CharCompatibleWithCodeset(CharAscii, CurrCodeSet) Then
            ' if we have a lookahead character AND if the next character is ALSO not compatible
            If (LookAheadAscii <> -1) AndAlso Not CharCompatibleWithCodeset(LookAheadAscii, CurrCodeSet) Then
                ' we need to switch code sets
                Select Case CurrCodeSet
                    Case CodeSet.CodeA
                        shifter = cCODEB
                        CurrCodeSet = CodeSet.CodeB
                        Exit Select
                    Case CodeSet.CodeB
                        shifter = cCODEA
                        CurrCodeSet = CodeSet.CodeA
                        Exit Select
                End Select
            Else
                ' no need to switch code sets, a temporary SHIFT will suffice
                shifter = cSHIFT
            End If
        End If

        If shifter <> -1 Then
            result = New Integer(1) {}
            result(0) = shifter
            result(1) = CodeValueForChar(CharAscii)
        Else
            result = New Integer(0) {}
            result(0) = CodeValueForChar(CharAscii)
        End If

        Return result
    End Function

    Public Shared Function CodesetAllowedForChar(CharAscii As Integer) As CodeSetAllowed
        If CharAscii >= 32 AndAlso CharAscii <= 95 Then
            Return CodeSetAllowed.CodeAorB
        Else
            Return If((CharAscii < 32), CodeSetAllowed.CodeA, CodeSetAllowed.CodeB)
        End If
    End Function

    Public Shared Function CharCompatibleWithCodeset(CharAscii As Integer, currcs As CodeSet) As Boolean
        Dim csa As CodeSetAllowed = CodesetAllowedForChar(CharAscii)
        Return csa = CodeSetAllowed.CodeAorB OrElse (csa = CodeSetAllowed.CodeA AndAlso currcs = CodeSet.CodeA) OrElse (csa = CodeSetAllowed.CodeB AndAlso currcs = CodeSet.CodeB)
    End Function

    Public Shared Function CodeValueForChar(CharAscii As Integer) As Integer
        Return If((CharAscii >= 32), CharAscii - 32, CharAscii + 64)
    End Function

    Public Shared Function StartCodeForCodeSet(cs As CodeSet) As Integer
        Return If(cs = CodeSet.CodeA, cSTARTA, cSTARTB)
    End Function

    Public Shared Function StopCode() As Integer
        Return cSTOP
    End Function

    Public Enum CodeSetAllowed
        CodeA
        CodeB
        CodeAorB
    End Enum
End Class