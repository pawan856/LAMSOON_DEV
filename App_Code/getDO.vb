Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization
Imports System.Data
Imports System.Data.SqlClient

Public Class IDOCodeResponse
    <XmlElement(ElementName:="DO_CODE")> _
   Public DO_CODE As String

    <XmlElement(ElementName:="STORER_CODE")> _
    Public STORER_CODE As String

    <XmlElement(ElementName:="DO_DELI_STATUS")> _
    Public DO_DELI_STATUS As String

    <XmlElement(ElementName:="DO_DEL_STATUS")> _
    Public DO_DEL_STATUS As String
End Class

Public Class IVehCodeResponse
    <XmlElement(ElementName:="VEH_LICENSE_PLATE")> _
   Public VEH_LICENSE_PLATE As String

    <XmlElement(ElementName:="VEH_NAME")> _
    Public VEH_NAME As String

    <XmlElement(ElementName:="VEH_REGION_CODE")> _
    Public VEH_REGION_CODE As String
End Class

Public Class IDOItemResponse
    <XmlElement(ElementName:="DOD_DISP_SEQ")> _
   Public DOD_DISP_SEQ As String

    <XmlElement(ElementName:="DOD_ITM_CODE")> _
    Public DOD_ITM_CODE As String

    <XmlElement(ElementName:="DOD_ITM_DESC")> _
    Public DOD_ITM_DESC As String

    <XmlElement(ElementName:="DOD_QTY")> _
    Public DOD_QTY As String

    <XmlElement(ElementName:="DOD_DELI_QTY")> _
    Public DOD_DELI_QTY As String

    <XmlElement(ElementName:="ITM_BARCODE")> _
    Public ITM_BARCODE As String

    <XmlElement(ElementName:="IMP_CODE")> _
    Public IMP_CODE As String
End Class

Public Class IDOResponse
    <XmlElement(ElementName:="DO_CODE")> _
   Public DO_CODE As String

    <XmlElement(ElementName:="STORER_CODE")> _
    Public STORER_CODE As String

    <XmlElement(ElementName:="DO_DELI_STATUS")> _
    Public DO_DELI_STATUS As String

    <XmlElement(ElementName:="DOD_DISP_SEQ")> _
    Public DOD_DISP_SEQ As String

    <XmlElement(ElementName:="DOD_ITM_CODE")> _
    Public DOD_ITM_CODE As String

    <XmlElement(ElementName:="DOD_ITM_DESC")> _
    Public DOD_ITM_DESC As String

    <XmlElement(ElementName:="DOD_QTY")> _
    Public DOD_QTY As String

    <XmlElement(ElementName:="DOD_DELI_QTY")> _
    Public DOD_DELI_QTY As String

    <XmlElement(ElementName:="ITM_BARCODE")> _
    Public ITM_BARCODE As String

    <XmlElement(ElementName:="IMP_CODE")> _
    Public IMP_CODE As String

    <XmlElement(ElementName:="DO_SIGNATURE")> _
    Public DO_SIGNATURE As String

    <XmlElement(ElementName:="DO_PRIORITY")> _
   Public DO_PRIORITY As String

    <XmlElement(ElementName:="DO_DEL_STATUS")> _
    Public DO_DEL_STATUS As String
End Class

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class getDO
    Inherits System.Web.Services.WebService

    Private gU As New GeneralUtils
    Private gDB As New GlobalDBFunc

    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function


    <WebMethod()> _
   Public Function getDOCodeList(ByVal vh_no As String) As IDOCodeResponse()
        Dim nSQL As String
        Dim datatbl As New DataTable

        nSQL = "SELECT DISTINCT WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.STORER_CODE, " & _
                    "WMS_DELV_ORDER.DO_STATUS,WMS_DELV_ORDER.DO_DELI_STATUS, " & _
                    "to_char(WMS_DELV_ORDER.DO_DATE, 'yy/mm/dd') as DO_DATE, " & _
                    "WMS_DELV_ORDER.DO_DEL_STATUS " & _
                    "FROM WMS_DELV_ORDER LEFT OUTER JOIN WMS_DELV_ORDER_D ON " & _
                    "WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND " & _
                    "WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND " & _
                    "WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE " & _
                    "LEFT OUTER JOIN WMS_ITEM ON " & _
                    "WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND " & _
                    "WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                    "WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY AND " & _
                    "WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
                    "WHERE WMS_DELV_ORDER.IMP_CODE = 'WMS'" & _
                    "AND WMS_DELV_ORDER.DO_SCH_DATE >= CONVERT(date, Getdate(), 111) " & _
                    "AND ISNULL(DO_VEHICLE_NO,'') <> ''"
        '"AND DO_VEHICLE_NO = N'" & vh_no & "' "


        Try
            datatbl = getDataTable(nSQL)

            Dim CodeResponse(datatbl.Rows.Count - 1) As IDOCodeResponse

            For i As Integer = 0 To datatbl.Rows.Count - 1
                CodeResponse(i) = New IDOCodeResponse
                CodeResponse(i).DO_CODE = datatbl.Rows(i).Item("DO_CODE").ToString
                CodeResponse(i).STORER_CODE = datatbl.Rows(i).Item("STORER_CODE").ToString
                CodeResponse(i).DO_DELI_STATUS = datatbl.Rows(i).Item("DO_DELI_STATUS").ToString
                CodeResponse(i).DO_DEL_STATUS = datatbl.Rows(i).Item("DO_DEL_STATUS").ToString
            Next

            Return CodeResponse

        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Function

    <WebMethod()> _
    Public Function getDOList(ByVal vh_no As String) As IDOResponse()
        Dim nSQL As String
        Dim datatbl As New DataTable

        nSQL = "SELECT WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.IMP_CODE, " & _
                    "WMS_DELV_ORDER.DO_STATUS,WMS_DELV_ORDER.DO_DELI_STATUS, " & _
                    "to_char(WMS_DELV_ORDER.DO_DATE, 'yy/mm/dd') as DO_DATE, " & _
                    "DOD_SEQ, DOD_DISP_SEQ, DOD_ITM_CODE, DOD_PACK_KEY, " & _
                    "DOD_ITM_DESC, DOD_QTY, DOD_DELI_QTY, ITM_BARCODE, " & _
                    "WMS_DELV_ORDER.DO_SIGNATURE, WMS_DELV_ORDER.DO_PRIORITY, " & _
                    "WMS_DELV_ORDER.DO_DEL_STATUS " & _
                    "FROM WMS_DELV_ORDER LEFT OUTER JOIN WMS_DELV_ORDER_D D ON " & _
                    "WMS_DELV_ORDER.IMP_CODE = D.IMP_CODE AND " & _
                    "WMS_DELV_ORDER.STORER_CODE = D.STORER_CODE AND " & _
                    "WMS_DELV_ORDER.DO_CODE = D.DO_CODE " & _
                    "LEFT OUTER JOIN WMS_ITEM ON " & _
                    "D.IMP_CODE = WMS_ITEM.IMP_CODE AND " & _
                    "D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                    "D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY AND " & _
                    "D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
                    "WHERE WMS_DELV_ORDER.IMP_CODE = 'WMS'" & _
                    "AND WMS_DELV_ORDER.DO_SCH_DATE >= CONVERT(date, Getdate(), 111) " & _
                    "AND ISNULL(WMS_DELV_ORDER.DO_VEHICLE_NO,'') <> '' "


        If vh_no <> "" Then
            nSQL += "AND DO_VEHICLE_NO = N'" & vh_no & "' "
        End If

        nSQL += "ORDER BY to_number(WMS_DELV_ORDER.DO_CODE) DESC, d.DOD_SEQ"
        '"AND DO_VEHICLE_NO = N'" & vh_no & "' "

        'nSQL = "SELECT DOD_SEQ, DOD_DISP_SEQ, DOD_ITM_CODE, DOD_PACK_KEY, " & _
        '       "DOD_ITM_DESC, DOD_QTY, DOD_DELI_QTY, ITM_BARCODE, d.STORER_CODE, d.IMP_CODE " & _
        '       "from WMS_DELV_ORDER_D d LEFT OUTER JOIN WMS_ITEM ON " & _
        '       "d.DOD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
        '       "where d.DO_CODE = '" & do_code & "' " & _
        '       "and d.IMP_CODE = 'WMS' " & _
        '       "order by d.DOD_DISP_SEQ"


        Try
            datatbl = getDataTable(nSQL)

            Dim CodeResponse(datatbl.Rows.Count - 1) As IDOResponse

            For i As Integer = 0 To datatbl.Rows.Count - 1
                CodeResponse(i) = New IDOResponse
                CodeResponse(i).DO_CODE = datatbl.Rows(i).Item("DO_CODE").ToString
                CodeResponse(i).STORER_CODE = datatbl.Rows(i).Item("STORER_CODE").ToString
                CodeResponse(i).DO_DELI_STATUS = datatbl.Rows(i).Item("DO_DELI_STATUS").ToString
                CodeResponse(i).DOD_DISP_SEQ = datatbl.Rows(i).Item("DOD_DISP_SEQ").ToString
                CodeResponse(i).DOD_ITM_CODE = datatbl.Rows(i).Item("DOD_ITM_CODE").ToString
                CodeResponse(i).DOD_ITM_DESC = datatbl.Rows(i).Item("DOD_ITM_DESC").ToString
                CodeResponse(i).DOD_QTY = datatbl.Rows(i).Item("DOD_QTY").ToString
                CodeResponse(i).DOD_DELI_QTY = datatbl.Rows(i).Item("DOD_DELI_QTY").ToString
                CodeResponse(i).ITM_BARCODE = datatbl.Rows(i).Item("ITM_BARCODE").ToString
                CodeResponse(i).IMP_CODE = datatbl.Rows(i).Item("IMP_CODE").ToString
                CodeResponse(i).DO_SIGNATURE = datatbl.Rows(i).Item("DO_SIGNATURE").ToString
                CodeResponse(i).DO_PRIORITY = datatbl.Rows(i).Item("DO_PRIORITY").ToString
                CodeResponse(i).DO_DEL_STATUS = datatbl.Rows(i).Item("DO_DEL_STATUS").ToString
            Next

            Return CodeResponse

        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Function


    <WebMethod()> _
    Public Function getVehList() As IVehCodeResponse()
        Dim nSQL As String
        Dim datatbl As New DataTable

        nSQL = "SELECT VEH_LICENSE_PLATE, VEH_LICENSE_PLATE+'('+VEH_REGION_CODE+')' as VEH_NAME, " & _
                "VEH_REGION_CODE FROM TMS_VEHICLE"


        Try
            datatbl = getDataTable(nSQL)

            Dim CodeResponse(datatbl.Rows.Count - 1) As IVehCodeResponse

            For i As Integer = 0 To datatbl.Rows.Count - 1
                CodeResponse(i) = New IVehCodeResponse
                CodeResponse(i).VEH_LICENSE_PLATE = datatbl.Rows(i).Item("VEH_LICENSE_PLATE").ToString
                CodeResponse(i).VEH_NAME = datatbl.Rows(i).Item("VEH_NAME").ToString
                CodeResponse(i).VEH_REGION_CODE = datatbl.Rows(i).Item("VEH_REGION_CODE").ToString
            Next

            Return CodeResponse

        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Function

    <WebMethod()> _
    Public Function getItemList(ByVal do_code As String) As IDOItemResponse()
        Dim nSQL As String
        Dim datatbl As New DataTable

        nSQL = "SELECT DOD_SEQ, DOD_DISP_SEQ, DOD_ITM_CODE, DOD_PACK_KEY, " & _
                   "DOD_ITM_DESC, DOD_QTY, DOD_DELI_QTY, ITM_BARCODE, d.STORER_CODE, d.IMP_CODE " & _
                   "from WMS_DELV_ORDER_D d LEFT OUTER JOIN WMS_ITEM ON " & _
                   "d.DOD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
                   "where d.DO_CODE = '" & do_code & "' " & _
                   "and d.IMP_CODE = 'WMS' " & _
                   "order by d.DOD_DISP_SEQ"

        Try
            datatbl = getDataTable(nSQL)

            Dim CodeResponse(datatbl.Rows.Count - 1) As IDOItemResponse

            For i As Integer = 0 To datatbl.Rows.Count - 1
                CodeResponse(i) = New IDOItemResponse
                CodeResponse(i).DOD_DISP_SEQ = datatbl.Rows(i).Item("DOD_DISP_SEQ").ToString
                CodeResponse(i).DOD_ITM_CODE = datatbl.Rows(i).Item("DOD_ITM_CODE").ToString
                CodeResponse(i).DOD_ITM_DESC = datatbl.Rows(i).Item("DOD_ITM_DESC").ToString
                CodeResponse(i).DOD_QTY = datatbl.Rows(i).Item("DOD_QTY").ToString
                CodeResponse(i).DOD_DELI_QTY = datatbl.Rows(i).Item("DOD_DELI_QTY").ToString
                CodeResponse(i).ITM_BARCODE = datatbl.Rows(i).Item("ITM_BARCODE").ToString
                CodeResponse(i).IMP_CODE = datatbl.Rows(i).Item("IMP_CODE").ToString
            Next

            Return CodeResponse

        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Function

    <WebMethod()> _
    Public Function saveItemList(ByVal doCode As String, ByVal StorerCode As String, _
                                    ByVal doSeq() As String, ByVal doQty() As String, ByVal doDeliQty() As String, _
                                    ByVal signature As String) As Boolean
        Dim datatbl As New DataTable
        Dim DO_DELI_STATUS As String = ""
        Dim checkConfirmFlag As Boolean = True
        Dim itemSQL As String = ""

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try


            For i As Integer = 0 To doSeq.Length - 1
                itemSQL = "update WMS_DELV_ORDER_D set " & _
                           "DOD_DELI_QTY = '" & gU.dbEncode(gU.decodeEmptyCInt(doDeliQty(i), 0)) & "', " & _
                           "sys_lub = N'MOBILE_DEVICE', " & _
                           "sys_lud = Getdate() " & _
                           "where IMP_CODE = 'WMS' " & _
                           "and STORER_CODE = '" & gU.dbEncode(StorerCode) & "' " & _
                           "and DO_CODE = '" & gU.dbEncode(doCode) & "' " & _
                           "and DOD_SEQ = '" & gU.dbEncode(doSeq(i)) & "'"

                If checkConfirmFlag Then
                    If gU.decodeEmptyCInt(doQty(i), 0) > gU.decodeEmptyCInt(doDeliQty(i), 0) Then
                        checkConfirmFlag = False
                    End If
                End If

                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
            Next


            If checkConfirmFlag = True Then
                DO_DELI_STATUS = "DELIVERED"
            Else
                DO_DELI_STATUS = "PARTIAL"
            End If

            Dim Sql As String = "update wms_delv_order " & _
                         "set do_deli_status = '" & DO_DELI_STATUS & "', " & _
                         "do_signature = N'" & signature & "', " & _
                         "sys_lub = N'MOBILE_DEVICE', " & _
                         "sys_lud = Getdate() " & _
                         "where imp_code = 'WMS' " & _
                         "and storer_code = '" & gU.dbEncode(StorerCode) & "' " & _
                         "and do_code = '" & gU.dbEncode(doCode) & "' "


            gDB.amendData(Sql)

            transaction.Commit()

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)
            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    <WebMethod()> _
    Public Function saveUCItemList(ByVal doCode As String, ByVal StorerCode As String, _
                                   ByVal doSeq() As String, ByVal doQty() As String, ByVal doDeliQty() As String, _
                                   ByVal signature As String, ByVal DO_DELI_STATUS As String, _
                                   ByVal DO_DEL_STATUS As String, ByVal fileList As String()) As Boolean
        Dim datatbl As New DataTable
        Dim itemSQL As String = ""

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            For i As Integer = 0 To doSeq.Length - 1
                itemSQL = "update WMS_DELV_ORDER_D set " & _
                           "DOD_DELI_QTY = '" & gU.dbEncode(gU.decodeEmptyCInt(doDeliQty(i), 0)) & "', " & _
                           "sys_lub = N'MOBILE_DEVICE', " & _
                           "sys_lud = Getdate() " & _
                           "where IMP_CODE = 'WMS' " & _
                           "and STORER_CODE = '" & gU.dbEncode(StorerCode) & "' " & _
                           "and DO_CODE = '" & gU.dbEncode(doCode) & "' " & _
                           "and DOD_SEQ = '" & gU.dbEncode(doSeq(i)) & "'"

                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
            Next

            Dim do_device_pic1, do_device_pic2, do_device_pic3, do_device_pic4, do_device_pic5 As Byte()

            'For i As Integer = 0 To fileList.Length - 1
            '    If fileList(i) IsNot Nothing Then
            '        Dim CompStr As New StrCompress(System.Text.Encoding.UTF8)

            '        CompStr.Passphrase = "tms"
            '        CompStr.Compressed = fileList(i).Trim

            '        Dim decryptFile As String

            '        decryptFile = CompStr.UnCompressed

            '        Dim picData As Byte() = Convert.FromBase64String(decryptFile)

            '        Select Case i
            '            Case 0
            '                do_device_pic1 = picData
            '            Case 1
            '                do_device_pic2 = picData
            '            Case 2
            '                do_device_pic3 = picData
            '            Case 3
            '                do_device_pic4 = picData
            '            Case 4
            '                do_device_pic5 = picData
            '        End Select
            '    End If


            '    'fileSQL += "DO_DEVICE_PIC" & (i + 1).ToString & "='" & fileList(i).Trim & "', "
            'Next


            'Dim Sql As String = "update wms_delv_order " & _
            '             "set do_deli_status = '" & DO_DELI_STATUS & "', " & _
            '             "do_del_status = N'" & DO_DEL_STATUS & "',  " & _
            '             "do_signature = N'" & signature & "', "

            Dim Sql As String = ""
            Sql += "update wms_delv_order "
            Sql += "set do_deli_status = @deliStatus, "
            Sql += "do_del_status = @delStatus,  "
            Sql += "do_signature = @sign, "
            Sql += "do_device_pic1 = @pic1, "
            Sql += "do_device_pic2 = @pic2, "
            Sql += "do_device_pic3 = @pic3, "
            Sql += "do_device_pic4 = @pic4, "
            Sql += "do_device_pic5 = @pic5, "


            'If fileSQL <> "" Then Sql += fileSQL

            Sql += "sys_lub = N'MOBILE_DEVICE', " & _
            "sys_lud = Getdate() " & _
            "where imp_code = 'WMS' " & _
            "and storer_code = '" & gU.dbEncode(StorerCode) & "' " & _
            "and do_code = '" & gU.dbEncode(doCode) & "' "


            Dim cmd As SqlCommand = New SqlCommand(Sql, gConn)

            Dim deliStatus As SqlParameter = cmd.Parameters.Add("@deliStatus", SqlDbType.VarChar, 10)
            Dim delStatus As SqlParameter = cmd.Parameters.Add("@delStatus", SqlDbType.VarChar, 100)
            Dim sign As SqlParameter = cmd.Parameters.Add("@sign", SqlDbType.VarChar, 4000)
            Dim pic1 As SqlParameter = cmd.Parameters.Add("@pic1", SqlDbType.VarBinary)
            Dim pic2 As SqlParameter = cmd.Parameters.Add("@pic2", SqlDbType.VarBinary)
            Dim pic3 As SqlParameter = cmd.Parameters.Add("@pic3", SqlDbType.VarBinary)
            Dim pic4 As SqlParameter = cmd.Parameters.Add("@pic4", SqlDbType.VarBinary)
            Dim pic5 As SqlParameter = cmd.Parameters.Add("@pic5", SqlDbType.VarBinary)


            deliStatus.Value = DO_DELI_STATUS
            delStatus.Value = DO_DEL_STATUS
            sign.Value = signature

            If do_device_pic1 IsNot Nothing Then pic1.Value = do_device_pic1 Else pic1.Value = DBNull.Value
            If do_device_pic2 IsNot Nothing Then pic2.Value = do_device_pic2 Else pic2.Value = DBNull.Value
            If do_device_pic3 IsNot Nothing Then pic3.Value = do_device_pic3 Else pic3.Value = DBNull.Value
            If do_device_pic4 IsNot Nothing Then pic4.Value = do_device_pic4 Else pic4.Value = DBNull.Value
            If do_device_pic5 IsNot Nothing Then pic5.Value = do_device_pic5 Else pic5.Value = DBNull.Value

            cmd.ExecuteNonQuery()


            'gDB.amendData(Sql)

            transaction.Commit()

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)
            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    <WebMethod()> _
    Public Function returnOrder(ByVal doCode As String, ByVal StorerCode As String) As Boolean
        Dim cancelSql As String = "update wms_delv_order " & _
                     "set do_deli_status = 'RETURN', " & _
                     "sys_lub = N'MOBILE_DEVICE', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = 'WMS' " & _
                     "and storer_code = '" & gU.dbEncode(StorerCode) & "' " & _
                     "and do_code = '" & gU.dbEncode(doCode) & "' "


        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)
            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    <WebMethod()> _
    Public Function SaveSignature(ByVal doCode As String, ByVal StorerCode As String, ByVal sign_data As String) As Boolean
        Dim cancelSql As String = "update wms_delv_order " & _
                     "set do_signature = N'" & sign_data & "', " & _
                     "sys_lub = N'MOBILE_DEVICE', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = 'WMS' " & _
                     "and storer_code = '" & gU.dbEncode(StorerCode) & "' " & _
                     "and do_code = '" & gU.dbEncode(doCode) & "' "


        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)
            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    Private Function getConnection() As SqlConnection
        Dim l_conn As SqlConnection

        l_conn = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)

        l_conn.Open()

        Return l_conn

    End Function

    Private Function getDataTable(ByVal as_sql As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing, Optional ByRef TableSchema As DataTable = Nothing) As DataTable
        Dim lds_1 As New DataSet
        Dim lda_1 As SqlDataAdapter
        Dim l_conn As SqlConnection
        Dim l_SqlCmd As SqlCommand

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If

        l_SqlCmd = New SqlCommand(as_sql, l_conn)
        If Not pTransaction Is Nothing Then
            l_SqlCmd.Transaction = pTransaction
        End If

        lda_1 = New SqlDataAdapter(l_SqlCmd)

        If Not TableSchema Is Nothing Then lda_1.FillSchema(TableSchema, SchemaType.Source)

        lda_1.Fill(lds_1)

        If pConn Is Nothing Then
            l_conn.Close()
            l_conn.Dispose()
        End If

        Return lds_1.Tables(0)

    End Function
End Class
