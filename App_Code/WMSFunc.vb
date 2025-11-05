Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Drawing

Public Class WMSFunc
    Private gU As New GeneralUtils
    Private gDB As New GlobalDBFunc
    Private db As New DBfunc

    Public Function getCartonQty(ByRef returnQty As Integer, ByVal fieldName As String, ByRef returnArray As ArrayList, _
                                 Optional ByVal getTotal As Boolean = False, _
                                 Optional ByVal d_row As DataRow = Nothing, _
                                 Optional ByVal d_table As DataTable = Nothing) As Boolean
        Dim cartonStr As String = ""
        Dim sArray As New ArrayList

        Try
            If getTotal Then
                Dim totalArray As New ArrayList

                If fieldName <> "" Then
                    If d_table IsNot Nothing Then

                        For i As Integer = 0 To d_table.Rows.Count - 1
                            cartonStr = ""
                            sArray = New ArrayList

                            cartonStr = d_table.Rows(i).Item(fieldName).ToString

                            sArray = SplitCarton(cartonStr)

                            If totalArray.Count = 0 Then
                                totalArray = sArray.Clone
                            Else
                                For x As Integer = 0 To sArray.Count - 1
                                    If Not totalArray.Contains(sArray(x)) Then
                                        totalArray.Add(sArray(x))
                                    End If
                                Next
                            End If
                        Next

                        returnQty = totalArray.Count

                        returnArray = totalArray

                        getCartonQty = True
                    End If
                End If
            Else
                If fieldName <> "" Then
                    If d_row IsNot Nothing Then
                        cartonStr = d_row(fieldName).ToString

                        sArray = SplitCarton(cartonStr)

                        returnQty = sArray.Count

                        returnArray = sArray

                        getCartonQty = True
                    End If
                End If
            End If
        Catch ex As Exception
            getCartonQty = False
        End Try
    End Function

    Public Function getCartonQtyFromGridView(ByRef returnQty As Integer, ByVal ControlName As String, _
                                             ByRef returnArray As ArrayList, Optional ByVal getTotal As Boolean = False, _
                                             Optional ByVal g_row As GridViewRow = Nothing, _
                                             Optional ByVal gv As GridView = Nothing) As Boolean

        Dim cartonStr As String = ""
        Dim sArray As New ArrayList

        Try
            If getTotal Then
                Dim totalArray As New ArrayList

                If ControlName <> "" Then
                    If gv IsNot Nothing Then

                        For i As Integer = 0 To gv.Rows.Count - 1
                            cartonStr = ""
                            sArray = New ArrayList

                            Dim ctl As Control = TryCast(gv.Rows(i).FindControl(ControlName), Control)

                            If ctl IsNot Nothing Then
                                If TypeOf ctl Is TextBox Then
                                    cartonStr = DirectCast(ctl, TextBox).Text
                                End If

                                sArray = SplitCarton(cartonStr)

                                If totalArray.Count = 0 Then
                                    totalArray = sArray.Clone
                                Else
                                    For x As Integer = 0 To sArray.Count - 1
                                        If Not totalArray.Contains(sArray(x)) Then
                                            totalArray.Add(sArray(x))
                                        End If
                                    Next
                                End If
                            End If
                        Next

                        returnQty = totalArray.Count

                        returnArray = totalArray

                        getCartonQtyFromGridView = True
                    End If
                End If
            Else
                If ControlName <> "" Then
                    If g_row IsNot Nothing Then
                        Dim ctl As Control = TryCast(g_row.FindControl(ControlName), Control)

                        If ctl IsNot Nothing Then
                            If TypeOf ctl Is TextBox Then
                                cartonStr = DirectCast(ctl, TextBox).Text
                            End If

                            sArray = SplitCarton(cartonStr)

                            returnQty = sArray.Count

                            returnArray = sArray

                            getCartonQtyFromGridView = True
                        End If

                    End If
                End If
            End If
        Catch ex As Exception
            getCartonQtyFromGridView = False
        End Try

    End Function

    Private Function SplitCarton(ByVal nStr As String) As ArrayList
        Dim nList() As String

        Dim isString As Boolean = False
        Dim numItem As Boolean = False
        Dim tempidx As Integer = -1

        Try
            SplitCarton = New ArrayList

            If nStr <> "" Then
                nList = Regex.Split(nStr & ",", "("")|(,)")

                Dim entry As String

                For Each entry In nList
                    If entry.Trim <> "" Then
                        If entry.Trim = "," Then

                            If tempidx >= 0 Then
                                SplitCarton.Add(tempidx)
                            End If

                            tempidx = -1
                            isString = False
                        Else
                            Dim x As Integer = 0

                            x = InStr(entry, "-")

                            If x > 0 Then

                                If entry.Substring(0, 1) <> "-" Then
                                    Dim toList() As String
                                    Dim fromNo As Integer = -1
                                    Dim toNo As Integer = -99

                                    toList = Split(entry, "-")

                                    For idx As Integer = 0 To toList.Length - 1
                                        If idx > 1 Then Exit For

                                        If IsNumeric(toList(idx)) Then
                                            If idx = 0 Then fromNo = gU.decodeEmptyCInt(toList(idx), -1)
                                            If idx = 1 Then toNo = gU.decodeEmptyCInt(toList(idx), -1)
                                        End If
                                    Next

                                    If fromNo <> toNo And fromNo < toNo Then
                                        For n As Integer = fromNo To toNo
                                            SplitCarton.Add(n)
                                        Next

                                    End If
                                End If

                                tempidx = -1

                                'tempstring = entry.Substring(0, x).Trim
                            Else
                                If IsNumeric(entry.Trim) Then
                                    tempidx = gU.decodeEmptyCInt(entry.Trim, -1)
                                End If
                            End If
                        End If
                    End If

                Next entry
            Else
                SplitCarton.Add("")
            End If

            Return SplitCarton

        Catch ex As Exception

            Throw ex
        End Try
    End Function

    Public Function GetDistanceByPoint(ByVal startPoint As Point, ByVal endPoint As Point) As Integer
        Return Math.Sqrt((Math.Abs(endPoint.X - startPoint.X) ^ 2) + _
                         (Math.Abs(endPoint.Y - startPoint.Y) ^ 2))
    End Function

    Public Function getSerialNo(ByVal Storer_code As String, ByVal itm_code As String, ByVal pack_key As String, ByVal ILBS_SERIAL_NO As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
        Dim returnSerial As String = ""
        Dim OriSerial As String = ""
        Dim tempSerial As String = ""
        Dim serialArr() As String
        Dim tempSeq As Integer = 0

        serialArr = Split(ILBS_SERIAL_NO, "-")

        If serialArr.Length = 2 Then
            OriSerial = serialArr(0)
        Else
            OriSerial = ILBS_SERIAL_NO
        End If

        Dim selectSQL As String = "Select ILBS_SERIAL_NO from WMS_ITEM_LOC_BAL_S where IMP_CODE='" & gU.dbEncode(HttpContext.Current.Session("IMP_CODE")) & "' and storer_code='" & Storer_code & "'" & _
                                  " and itm_code='" & gU.dbEncode(itm_code) & "' and pack_key='" & gU.dbEncode(pack_key) & "' and UPPER(ILBS_SERIAL_NO)=UPPER('" & gU.dbEncode(OriSerial) & "')"

        tempSerial = db.getValueFromSQL(selectSQL, pConn, pTransaction)

        If tempSerial = "" Then
            returnSerial = OriSerial
        Else
            selectSQL = "Select max(ILBS_SERIAL_NO) from from WMS_ITEM_LOC_BAL_S where IMP_CODE='" & gU.dbEncode(HttpContext.Current.Session("IMP_CODE")) & "' and storer_code='" & Storer_code & "'" & _
                        " and itm_code='" & gU.dbEncode(itm_code) & "' and pack_key='" & gU.dbEncode(pack_key) & "' and UPPER(ILBS_SERIAL_NO) like UPPER('" & gU.dbEncode(OriSerial) & "%')"

            tempSerial = db.getValueFromSQL(selectSQL, pConn, pTransaction)

            serialArr = Split(tempSerial, "-")

            If serialArr.Length > 1 Then
                If IsNumeric(serialArr(serialArr.Length - 1)) Then
                    tempSeq = CInt(serialArr(serialArr.Length - 1)) + 1
                Else
                    tempSeq = 1
                End If

                returnSerial = tempSerial & "-" & tempSeq
            Else
                returnSerial = tempSerial & "-1"
            End If

        End If

        Return returnSerial
    End Function
End Class
