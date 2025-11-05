Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports System.Globalization
Imports ExcelLibrary.SpreadSheet
Imports System.IO
Imports System.Net.Mail

Public Class cronjobImplement
    Inherits System.Web.UI.Page
    Private moduleAction As String = ""
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private impExp As New IMPORTEXPORT
    Private st As New StockTrans
    Private gU As New GeneralUtils
    Private CurrTrans As String = ""

    Public Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session("usr_Id") = "0"
        Session("usr_nickname") = "CronJob"
        Session("gLang") = ""
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim cronType As String
        cronType = ""
        If Not (Request.QueryString("crontype") Is Nothing) Then
            If Request.QueryString("crontype").ToString() <> "" Then
                cronType = Request.QueryString("crontype").ToString
            End If
        End If
        If cronType.ToLower = "download" Then
            'For Download
            ImportSOADJData()
            ImportPOData()
            ImportTOData()
            ImportIPRData()
            ImportSRData()
            ImportSOData()
            SendWrongLOTEmail()
        ElseIf cronType.ToLower = "upload" Then
            'For Upload
            ImportINV_TRANSFERData()
            ImportPO_GRNData()
            ImportTO_GRNData()
            ImportIPR_GRNData()
            ImportSO_GRNData()
            ImportSR_GRNData()
        ElseIf cronType.ToLower = "master" Then
            'For Master
            ImportUOMMASTERData()
            ImportWAREHOUSEMASTERData()
            ImportITEMMASTERData()
            ImportWAREHOUSELOCATIONData()
            ImportPACKINGRULEData()
            'CloseSQLConnection()
        ElseIf cronType.ToLower = "custmaster" Then
            'For Customer Master
            ImportCUSTOMERMASTERData()
        ElseIf cronType.ToLower = "matchreport" Then
            'For Matching report email
            SendMailMatchingOutRpt()
        ElseIf cronType.ToLower = "sodownload" Then
            'For Importing SO Data
            ImportSOData()
        ElseIf cronType.ToLower = "archivedata" Then
            'For Archiving Data
            ArchiveData()
        ElseIf cronType.ToLower = "reindexing" Then
            'For Archiving Data
            ReIndexing()
        ElseIf cronType.ToLower = "generatebackup" Then
            'For Backup Data
            GenerateBackup()
        ElseIf cronType.ToLower = "delfiles" Then
            DelTempDir()
        End If


    End Sub

    'Import STOCK ADJUSTMENT
    Public Sub ImportSOADJData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        'Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        'Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim currTrans As String = ""
        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now
            Dim STKAdjCodes As String = ""
            SQLString = "Select ACTION,DOC_TYPE,TRANSACTION_ID,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS = 'NEW' and DOC_TYPE='ADJ'"

            Dim codtl As New DataTable
            codtl = gDB.getDataTable(SQLString, gConn, transaction)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim SOADJInsert As Int16 = 0
                currTrans = ""
                For Each rowD As DataRow In codtl.Rows
                    currTrans = rowD("TRANSACTION_ID").ToString.Trim
                    If (rowD("ACTION").ToString.Trim) = "NEW" And (rowD("DOC_TYPE").ToString.Trim) = "ADJ" Then
                        SQLString = "Select * from EBS_WMS_TRANS_ITX_STOCK_ADJUSTMENT WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_STOCK_ADJUST " &
                    "WHERE AD_REF_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode Is Nothing Or dtROACTIONCode.Rows.Count <= 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    'For STORER 
                                    Dim IO_CODE As String
                                    SQLString = "select Top(1) IO_ID from EBS_WMS_COMPANY_MASTER where IO_CODE = '" & rowH("IO_CODE").ToString.Trim & "' "
                                    Dim dtIOCode As DataTable
                                    dtIOCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtIOCode IsNot Nothing AndAlso dtIOCode.Rows.Count > 0 Then
                                        IO_CODE = dtIOCode.Rows(0)("IO_ID").ToString.Trim
                                    Else
                                        IO_CODE = ""
                                    End If

                                    'for WMS_STOCK_ADJUST '
                                    SQLString = "Select * from WMS_STOCK_ADJUST " &
                      "WHERE AD_REF_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                                        'INSERT
                                        nextNo = DB.getDocNo("SADJ", gConn, transaction)
                                        Session("ADJnxtNo") = nextNo
                                        SQLString = "Insert into WMS_STOCK_ADJUST(IMP_CODE,STORER_CODE,AD_CODE,AD_DATE,AD_REF_CK_CODE,AD_TYPE,AD_STATUS,AD_WH,AD_REF_NO,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE, @STORER_CODE,@AD_CODE,@AD_DATE,@AD_REF_CK_CODE,@AD_TYPE,@AD_STATUS,@AD_WH,@AD_REF_NO,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                        cmd.Parameters.AddWithValue("@STORER_CODE", IO_CODE)
                                        cmd.Parameters.AddWithValue("@AD_CODE", Session("ADJnxtNo"))
                                        cmd.Parameters.AddWithValue("@AD_TYPE", "EBS")
                                        cmd.Parameters.AddWithValue("@AD_DATE", rowH("TRANSACTION_DATE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@AD_STATUS", "New")
                                        cmd.Parameters.AddWithValue("@AD_WH", rowH("SUBINVENTORY_CODE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@AD_REF_CK_CODE", rowH("BATCH_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@AD_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now) 'rowH("TRANSACTION_DATE").ToString.Trim
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        insertCount = insertCount + 1

                                        'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION Set STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                        SOADJInsert = 1

                                        'for WMS_STOCK_ADJUST details'
                                        SQLString = "Select * from EBS_WMS_TRANS_ITX_STOCK_ADJUSTMENT Where TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "' "
                                        'TO_SUBINVENTORY_CODE
                                        Dim codtTemp As New DataTable
                                        codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                        If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                            For Each rowDtl As DataRow In codtTemp.Rows

                                                Dim AdSeq As String
                                                SQLString = "Select IsNUll(Count(AD_CODE),0)+1 AS AD_SEQ from WMS_STOCK_ADJUST_D where AD_CODE='" + Session("ADJnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & IO_CODE & "'"
                                                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                    AdSeq = 0
                                                Else
                                                    AdSeq = dtRPSD.Rows(0)("AD_SEQ").ToString
                                                End If

                                                'For ITM_CODE
                                                Dim ITM_CODE As String
                                                SQLString = "select Top(1) ITM_CODE as  ITEM_ID from WMS_ITEM where ITM_SKU_NO = '" & rowDtl("ITEM_NUMBER").ToString.Trim & "' "
                                                Dim dtITMOCode As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtITMOCode IsNot Nothing AndAlso dtITMOCode.Rows.Count > 0 Then
                                                    ITM_CODE = dtITMOCode.Rows(0)("ITEM_ID").ToString
                                                Else
                                                    ITM_CODE = ""
                                                End If

                                                'For location
                                                Dim STKLOC As String = rowDtl("LOCATOR").ToString.Trim.Replace(".", "")

                                                'For ORG_QTY 
                                                Dim ORG_QTY As String
                                                SQLString = "select Top(1) ILOC_BAL_QTY from WMS_ITEM_LOC_BAL where ITM_CODE = '" & ITM_CODE & "' and ILOC_BATCH_NO='" & rowDtl("LOT_NUMBER").ToString.Trim & "' and ILOC_LOC='" & STKLOC & "'"
                                                Dim dtORGQTY As DataTable
                                                dtORGQTY = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtORGQTY IsNot Nothing AndAlso dtORGQTY.Rows.Count > 0 Then
                                                    ORG_QTY = dtORGQTY.Rows(0)("ILOC_BAL_QTY").ToString.Trim
                                                Else
                                                    ORG_QTY = 0
                                                End If

                                                Dim Expiry_Date As String
                                                SQLString = "select top(1) ILOC_EXPIRY_DATE from WMS_ITEM_LOC_BAL where STORER_CODE=" & IO_CODE & " and IMP_CODE='WMS' and ITM_CODE='" & ITM_CODE & "' and iloc_batch_no='" & rowDtl("LOT_NUMBER").ToString.Trim & "'  and ILOC_WH='" & rowDtl("SUBINVENTORY_CODE").ToString.Trim & "' and ILOC_LOC='" & rowDtl("LOCATOR").ToString.Trim.Replace(".", "") & "' order by ILOC_EXPIRY_DATE  desc "
                                                'SQLString = "Select TOP(1) EXPIRATION_DATE from EBS_WMS_STOCK_ONHAND where IO_CODE=" & IO_CODE & " and SUBINVENTORY_CODE='" & rowDtl("SUBINVENTORY_CODE").ToString.Trim & "' and LOT_NUMBER='" & rowDtl("LOT_NUMBER").ToString.Trim & "' and ITEM_NUMBER='" & rowDtl("ITEM_NUMBER").ToString.Trim & "' and LOCATOR='" & rowDtl("LOCATOR").ToString.Trim & "'"
                                                Dim dtExpCode As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtExpCode IsNot Nothing AndAlso dtExpCode.Rows.Count > 0 Then
                                                    Expiry_Date = dtExpCode.Rows(0)("ILOC_EXPIRY_DATE").ToString
                                                Else
                                                    Expiry_Date = ""
                                                End If

                                                SQLString = "Select * from WMS_STOCK_ADJUST_D " &
                             "WHERE ADD_SERIAL_NO ='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and DRUM_LEVEL = '" & (rowDtl("BATCH_NO").ToString.Trim) & "' and DRUM_ID = '" & (rowDtl("SEQ_NO").ToString.Trim) & "' "

                                                Dim codtlRep As DataTable
                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                    'INSERT
                                                    'Parameterized Parameter
                                                    SQLString = "Insert into WMS_STOCK_ADJUST_D(IMP_CODE,STORER_CODE,AD_CODE,AD_SEQ,ADD_PACK_KEY,ADD_ORG_QTY,ADD_REV_QTY,ADD_SERIAL_NO,ADD_EXPIRY_DATE,DRUM_ID,DRUM_LEVEL,ADD_VND_CODE,ADD_ITM_CODE,ADD_VAR_QTY,ADD_BATCH_NO,ADD_LOC,ADD_PALLET_NO,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD,ADD_MANU_DATE) VALUES (@IMP_CODE,@STORER_CODE,@AD_CODE,@AD_SEQ,@ADD_PACK_KEY,@ADD_ORG_QTY,@ADD_REV_QTY,@ADD_SERIAL_NO,@ADD_EXPIRY_DATE,@DRUM_ID,@DRUM_LEVEL,@ADD_VND_CODE,@ADD_ITM_CODE,@ADD_VAR_QTY,@ADD_BATCH_NO,@ADD_LOC,@ADD_PALLET_NO,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD,@ADD_MANU_DATE)"
                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                    cmd.Parameters.AddWithValue("@STORER_CODE", IO_CODE)
                                                    cmd.Parameters.AddWithValue("@AD_CODE", Session("ADJnxtNo"))
                                                    cmd.Parameters.AddWithValue("@AD_SEQ", AdSeq)
                                                    cmd.Parameters.AddWithValue("@ADD_PACK_KEY", "1")
                                                    cmd.Parameters.AddWithValue("@ADD_ITM_CODE", ITM_CODE)
                                                    If rowDtl("TRANSACTION_TYPE").ToString.Trim = "I" Then
                                                        cmd.Parameters.AddWithValue("@ADD_VAR_QTY", Convert.ToDecimal(rowDtl("QUANTITY").ToString.Trim) * -1.0)
                                                    ElseIf rowDtl("TRANSACTION_TYPE").ToString.Trim = "R" Then
                                                        cmd.Parameters.AddWithValue("@ADD_VAR_QTY", rowDtl("QUANTITY").ToString.Trim)
                                                    End If
                                                    cmd.Parameters.AddWithValue("@ADD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ADD_LOC", rowDtl("LOCATOR").ToString.Trim.Replace(".", ""))
                                                    cmd.Parameters.AddWithValue("@ADD_PALLET_NO", rowDtl("SUBINVENTORY_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ADD_ORG_QTY", ORG_QTY)
                                                    If rowDtl("TRANSACTION_TYPE").ToString.Trim = "I" Then
                                                        cmd.Parameters.AddWithValue("@ADD_REV_QTY", (ORG_QTY - rowDtl("QUANTITY")))
                                                    ElseIf rowDtl("TRANSACTION_TYPE").ToString.Trim = "R" Then
                                                        cmd.Parameters.AddWithValue("@ADD_REV_QTY", (ORG_QTY + rowDtl("QUANTITY")))
                                                    End If
                                                    cmd.Parameters.AddWithValue("@ADD_SERIAL_NO", DBNull.Value)
                                                    cmd.Parameters.AddWithValue("@ADD_EXPIRY_DATE", Expiry_Date)
                                                    cmd.Parameters.AddWithValue("@DRUM_ID", rowDtl("SEQ_NO").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@DRUM_LEVEL", rowDtl("BATCH_NO").ToString.Trim.Substring("0", "8"))
                                                    cmd.Parameters.AddWithValue("@ADD_VND_CODE", rowDtl("REQUEST_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                    cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                                    cmd.Parameters.AddWithValue("@SYS_CD", rowDtl("TRANSACTION_DATE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@SYS_LUD", rowDtl("TRANSACTION_DATE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ADD_MANU_DATE", IIf(String.IsNullOrEmpty(rowDtl("EXPIRY_DATE").ToString.Trim), DBNull.Value, rowDtl("EXPIRY_DATE").ToString.Trim))
                                                    cmd.CommandType = System.Data.CommandType.Text
                                                    cmd.ExecuteScalar()
                                                    insertCountD = insertCountD + 1
                                                End If

                                            Next

                                        End If

                                        STKAdjCodes += Session("ADJnxtNo") & ","


                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql, gConn, transaction)

                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)
                            lblMSG.Text = "STOCK ADJUSTMENT data has been already updated"

                        End If

                    End If

                Next
                currTrans = ""

                'For INTERFACE LOG
                If SOADJInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "SOADJ_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_SOADJ_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_STOCK_ADJUST"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                            "WHERE ACTION <> 'CLOSED' and DOC_TYPE='ADJ'"
                    Dim ADJITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If ADJITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If

                End If

            Else
                lblMSG.Text = "STOCK ADJUSTMENT data has been already updated"
            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>STOCK ADJ Inserted Header Row : " & insertCount.ToString.Trim & "</p>"
                'strAlert &= "<p>STOCK ADJ Updated Header Row : " & updateCount.ToString.Trim & "</p>"
                strAlert &= "<p>STOCK ADJ Inserted Details Row : " & insertCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>STOCK ADJ Updated Details Row : " & updateCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>STOCK ADJ data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "STOCK ADJUSTMENT data has been already updated"
            End If

            'Post downloaded Stock Adjustments
            POSTDATASTKADJ(STKAdjCodes)

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='" + ex.Message.Replace("'", "") + "',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & currTrans & "'"
            gDB.amendData(UpdateSql)
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import PO
    Public Sub ImportPOData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select ACTION,DOC_TYPE,DOC_ID,TRANSACTION_ID,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS = 'NEW' and DOC_TYPE='PO' ORDER BY TRANSACTION_ID ASC"

            Dim codtl As New DataTable
            codtl = gDB.getDataTable(SQLString, gConn, transaction)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim POInsert As Int16 = 0

                For Each rowD As DataRow In codtl.Rows
                    If ((rowD("ACTION").ToString.Trim) = "NEW" Or (rowD("ACTION").ToString.Trim = "UPDATE")) And (rowD("DOC_TYPE").ToString.Trim) = "PO" Then

                        'INSERT
                        SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,IO_ID,PO_NO,VENDOR_NAME from EBS_WMS_TRANS_ITX_PO_HEADER where TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then

                            ' Delete existing PO data if action is update and NEW PO exists
                            If rowD("ACTION").ToString.Trim = "UPDATE" Then

                                SQLString = "select 1 from WMS_GOODSRCV where GR_DOC_NO=(select RO_CODE from WMS_REPLENISH where RO_REF_NO='" & codt.Rows(0).Item("PO_HEADER_ID") & "')"
                                Dim dtCheck As New DataTable
                                dtCheck = gDB.getDataTable(SQLString, gConn, transaction)

                                UpdateSql = " if not exists (select 1 from WMS_GOODSRCV where GR_DOC_NO=(select RO_CODE from WMS_REPLENISH where RO_REF_NO='" & codt.Rows(0).Item("PO_HEADER_ID") & "'))  " &
                                            " begin  " &
                                            " Delete from WMS_REPLENISH_D where RO_CODE = (select RO_CODE from WMS_REPLENISH where RO_REF_NO='" & codt.Rows(0).Item("PO_HEADER_ID") & "')  " &
                                            " Delete from WMS_REPLENISH where RO_CODE = (select RO_CODE from WMS_REPLENISH where RO_REF_NO='" & codt.Rows(0).Item("PO_HEADER_ID") & "')  " &
                                            " end  " &
                                            " else  " &
                                            " begin  " &
                                            " UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='GR Already Done',  " &
                                            " LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'  " &
                                            " end  "
                                gDB.amendData(UpdateSql, gConn, transaction)

                                If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                                    Continue For
                                End If

                            End If

                            SQLString = "Select * from WMS_REPLENISH " &
                    "WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode Is Nothing Or dtROACTIONCode.Rows.Count <= 0 Then

                                For Each rowH As DataRow In codt.Rows
                                    'For PO_CAT
                                    Dim POCat As String
                                    SQLString = "Select IsNull(STO_CAT,'') AS PO_CAT from WMS_STORER where STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "'"
                                    Dim dtCAT As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtCAT Is Nothing Or dtCAT.Rows.Count <= 0 Then
                                        POCat = ""
                                    Else
                                        POCat = dtCAT.Rows(0)("PO_CAT").ToString.Trim
                                    End If

                                    'for PO DATE'
                                    Dim PODATE As String
                                    SQLString = "Select CREATION_DATE from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='PO' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString.Trim & "'"
                                    Dim dtPODATE As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtPODATE Is Nothing Or dtPODATE.Rows.Count <= 0 Then
                                        PODATE = ""
                                    Else
                                        PODATE = dtPODATE.Rows(0)("CREATION_DATE").ToString
                                    End If


                                    SQLString = "Select * from WMS_REPLENISH " &
                   "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                                        'INSERT
                                        nextNo = DB.getDocNo("RO", gConn, transaction)
                                        Session("ROnxtNo") = nextNo
                                        SQLString = "Insert into WMS_REPLENISH(IMP_CODE,STORER_CODE,RO_CODE,RO_DATE,RO_EDI_PO_NO,RO_SEAL_NO,RO_REF_NO,RO_STATUS,PO_TYPE,PO_CAT,RO_TYPE,RO_ISSUED_BY,RO_WH_CODE,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE, @STORER_CODE,@RO_CODE,@RO_DATE,@RO_EDI_PO_NO,@RO_SEAL_NO,@RO_REF_NO,@RO_STATUS,@PO_TYPE,@PO_CAT,@RO_TYPE,@RO_ISSUED_BY,@RO_WH_CODE,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                        cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_CODE", Session("ROnxtNo"))
                                        cmd.Parameters.AddWithValue("@RO_DATE", PODATE)
                                        cmd.Parameters.AddWithValue("@RO_EDI_PO_NO", rowH("PO_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("PO_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_STATUS", "NEW")
                                        cmd.Parameters.AddWithValue("@PO_TYPE", rowD("DOC_TYPE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_TYPE", rowD("DOC_TYPE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@PO_CAT", POCat)
                                        cmd.Parameters.AddWithValue("@RO_ISSUED_BY", rowH("VENDOR_NAME").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_WH_CODE", "FG01")
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        insertCount = insertCount + 1

                                        'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql, gConn, transaction)

                                        POInsert = 1

                                        SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,PO_DISTRIBUTION_ID,IO_ID,LINE_NUM,ITEM_ID,ITEM_NUMBER,ITEM_DESCRIPTION,UOM_CODE,QUANTITY_ORDERED from EBS_WMS_TRANS_ITX_PO_DETAIL Where PO_HEADER_ID='" & rowH("PO_HEADER_ID").ToString() & "' and  TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "'"
                                        Dim codtTemp As New DataTable
                                        codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                        If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then
                                            For Each rowDtl As DataRow In codtTemp.Rows
                                                Dim RdSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("ROnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                    RdSeq = 0
                                                Else
                                                    RdSeq = dtRPSD.Rows(0)("ROD_SEQ").ToString
                                                End If

                                                'FOR SNO.'
                                                Dim RdDispSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_DISP_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("ROnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                Dim dtDISP As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtDISP Is Nothing Or dtDISP.Rows.Count <= 0 Then
                                                    RdDispSeq = 0
                                                Else
                                                    RdDispSeq = dtDISP.Rows(0)("ROD_DISP_SEQ").ToString
                                                End If

                                                'for ITEM NAME'
                                                SQLString = "select Top(1) ITM_NAME ITEM_DESCRIPTION from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                                                Dim dtITM As DataTable
                                                dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                                                Dim ITMName As String
                                                If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                    ITMName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                                Else
                                                    ITMName = ""
                                                End If

                                                SQLString = "Select * from WMS_REPLENISH_D " &
                                 "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("PO_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO='" & (rowDtl("PO_DISTRIBUTION_ID").ToString.Trim) & "'"
                                                Dim codtlRep As DataTable
                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                    'INSERT
                                                    'Parameterized Parameter
                                                    SQLString = "Insert into WMS_REPLENISH_D(IMP_CODE,STORER_CODE,ROD_SEQ,ROD_DISP_SEQ,RO_CODE,ROD_ITM_CODE,ROD_ITM_NAME,ROD_SKU_NO,ROD_UOM,ROD_QTY,ROD_REF_NO,ROD_DOC_NO,ROD_SERIES_NO,ROD_PALLET_NO,ROD_STATUS,ROD_PACK_KEY,ROD_WH_CODE,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE,@STORER_CODE,@ROD_SEQ,@ROD_DISP_SEQ,@RO_CODE,@ROD_ITM_CODE,@ROD_ITM_NAME,@ROD_SKU_NO,@ROD_UOM,@ROD_QTY,@ROD_REF_NO,@ROD_DOC_NO,@ROD_SERIES_NO,@ROD_PALLET_NO,@ROD_STATUS,@ROD_PACK_KEY,@ROD_WH_CODE,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SEQ", RdSeq)
                                                    cmd.Parameters.AddWithValue("@ROD_DISP_SEQ", RdDispSeq)
                                                    cmd.Parameters.AddWithValue("@RO_CODE", Session("ROnxtNo"))
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_NAME", ITMName)

                                                    cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowDtl("ITEM_NUMBER").ToString.Trim)

                                                    cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("UOM_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY_ORDERED").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("LINE_NUM").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("PO_HEADER_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("PO_DISTRIBUTION_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_STATUS", "NEW")
                                                    cmd.Parameters.AddWithValue("@ROD_PACK_KEY", "1")
                                                    cmd.Parameters.AddWithValue("@ROD_WH_CODE", "FG01")
                                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                    cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                                    cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                                    cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                                    cmd.CommandType = System.Data.CommandType.Text
                                                    cmd.ExecuteScalar()
                                                    insertCountD = insertCountD + 1
                                                End If

                                            Next

                                        End If


                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE STATUS ! = 'ERROR' AND TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE  STATUS ! = 'ERROR' AND TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)
                            lblMSG.Text = "PO data has been already updated"

                        End If

                        'ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "PO" Then
                        '    'UPDATE
                        '    SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,IO_ID,PO_NO,VENDOR_NAME from EBS_WMS_TRANS_ITX_PO_HEADER WHERE PO_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "' "
                        '    'SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,IO_ID,PO_NO,VENDOR_NAME from EBS_WMS_TRANS_ITX_PO_HEADER WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        '    Dim codt As New DataTable
                        '    codt = gDB.getDataTable(SQLString, gConn, transaction)

                        '    If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                        '        SQLString = "Select * from WMS_REPLENISH " &
                        '    "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        '        Dim dtROACTIONCode As DataTable
                        '        dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                        '        If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then
                        '            For Each rowH As DataRow In codt.Rows
                        '                SQLString = "Select * from WMS_REPLENISH " &
                        '    "WHERE  STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                        '                Dim dtROCode As DataTable
                        '                dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                        '                If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                        '                    sbCmdText = New StringBuilder()
                        '                    sbCmdText.Append("Update WMS_REPLENISH Set ")
                        '                    sbCmdText.Append("RO_EDI_PO_NO = @RO_EDI_PO_NO,")
                        '                    'sbCmdText.Append("RO_STATUS = @RO_STATUS,")
                        '                    sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        '                    sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '                    sbCmdText.Append(" Where RO_REF_NO=@RO_REF_NO and STORER_CODE=@STORER_CODE")
                        '                    cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        '                    cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RO_EDI_PO_NO", rowH("PO_NO").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("PO_HEADER_ID").ToString.Trim)
                        '                    'cmd.Parameters.AddWithValue("@RO_STATUS", "NEW")
                        '                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        '                    cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        '                    cmd.CommandType = System.Data.CommandType.Text
                        '                    cmd.ExecuteScalar()
                        '                    updateCount = updateCount + 1

                        '                    'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                        '                    UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                        '                    gDB.amendData(UpdateSql, gConn, transaction)


                        '                    SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,PO_DISTRIBUTION_ID,IO_ID,LINE_NUM,ITEM_ID,ITEM_NUMBER,ITEM_DESCRIPTION,UOM_CODE,QUANTITY_ORDERED from EBS_WMS_TRANS_ITX_PO_DETAIL Where PO_HEADER_ID='" & rowH("PO_HEADER_ID").ToString() & "' and  TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "'"
                        '                    Dim codtTemp As New DataTable
                        '                    codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                        '                    If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                        '                        For Each rowDtl As DataRow In codtTemp.Rows
                        '                            'for ITEM NAME'
                        '                            SQLString = "select Top(1) ITM_NAME ITEM_DESCRIPTION from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                        '                            Dim dtITM As DataTable
                        '                            dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                        '                            Dim ITMName As String
                        '                            If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                        '                                ITMName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                        '                            Else
                        '                                ITMName = ""
                        '                            End If

                        '                            SQLString = "Select * from WMS_REPLENISH_D " &
                        '             "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("PO_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO='" & (rowDtl("PO_DISTRIBUTION_ID").ToString.Trim) & "'and ROD_REF_NO='" & (rowDtl("LINE_NUM").ToString.Trim) & "'"
                        '                            Dim codtlRep As DataTable
                        '                            codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                        '                            If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                        '                                'UPDATE
                        '                                sbCmdText = New StringBuilder()
                        '                                sbCmdText.Append("Update WMS_REPLENISH_D Set ")
                        '                                sbCmdText.Append("ROD_ITM_NAME = @ROD_ITM_NAME,")
                        '                                sbCmdText.Append("ROD_ITM_CODE = @ROD_ITM_CODE,")
                        '                                sbCmdText.Append("ROD_SKU_NO = @ROD_SKU_NO,")
                        '                                sbCmdText.Append("ROD_UOM = @ROD_UOM,")
                        '                                sbCmdText.Append("ROD_QTY = @ROD_QTY,")
                        '                                sbCmdText.Append("ROD_REF_NO = @ROD_REF_NO,")
                        '                                'sbCmdText.Append("ROD_STATUS = @ROD_STATUS,")
                        '                                sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        '                                sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '                                sbCmdText.Append(" Where ROD_DOC_NO=@ROD_DOC_NO and ROD_SERIES_NO=@ROD_SERIES_NO and ROD_PALLET_NO=@ROD_PALLET_NO")
                        '                                cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        '                                cmd.Parameters.AddWithValue("@ROD_ITM_NAME", ITMName)
                        '                                cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowDtl("ITEM_NUMBER").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("UOM_CODE").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY_ORDERED").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("LINE_NUM").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("PO_HEADER_ID").ToString.Trim)
                        '                                cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("PO_DISTRIBUTION_ID").ToString.Trim)
                        '                                'cmd.Parameters.AddWithValue("@ROD_STATUS", "NEW")
                        '                                cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        '                                cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        '                                cmd.CommandType = System.Data.CommandType.Text
                        '                                cmd.ExecuteScalar()
                        '                                updateCountD = updateCountD + 1
                        '                            End If

                        '                        Next

                        '                    End If

                        '                End If

                        '            Next

                        '        Else
                        '            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        '            gDB.amendData(UpdateSql, gConn, transaction)
                        '        End If

                        '    Else
                        '        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        '        gDB.amendData(UpdateSql, gConn, transaction)
                        '    End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "PO" Then
                        SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,IO_ID,PO_NO,VENDOR_NAME from EBS_WMS_TRANS_ITX_PO_HEADER WHERE PO_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                            "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                            "WHERE RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CANCEL' WHERE RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and 
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("PO_HEADER_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CLOSED" And (rowD("DOC_TYPE").ToString.Trim) = "PO" Then
                        SQLString = "Select TRANSACTION_ID,PO_HEADER_ID,IO_ID,PO_NO,VENDOR_NAME from EBS_WMS_TRANS_ITX_PO_HEADER WHERE PO_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and 
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CLOSED' WHERE RO_REF_NO='" & (rowH("PO_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and 
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("PO_HEADER_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    End If

                Next

                'For INTERFACE LOG
                If POInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "PO_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_PO_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_REPLENISH WHERE PO_TYPE='PO'"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='PO'"
                    Dim POITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If POITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If

                End If

            Else
                lblMSG.Text = "PO data has been already updated"
            End If


            'FOR NPO'
            SQLString = "SELECT rp.*,gr.* FROM WMS_GOODSRCV gr INNER JOIN WMS_REPLENISH rp ON rp.RO_EDI_PO_NO=gr.VND_CODE and rp.STORER_CODE=gr.STORER_CODE"
            Dim codtlNPO As New DataTable
            codtlNPO = gDB.getDataTable(SQLString, gConn, transaction)
            If codtlNPO IsNot Nothing AndAlso codtlNPO.Rows.Count > 0 Then
                For Each row As DataRow In codtlNPO.Rows
                    sbCmdText = New StringBuilder()
                    sbCmdText.Append("Update WMS_GOODSRCV Set ")
                    sbCmdText.Append("GR_TYPE = @GR_TYPE,")
                    sbCmdText.Append("PO_CODE = @PO_CODE,")
                    sbCmdText.Append("GR_DOC_TYPE = @GR_DOC_TYPE,")
                    sbCmdText.Append("GR_DOC_NO = @GR_DOC_NO,")
                    sbCmdText.Append("GR_EDI_PO_NO = @GR_EDI_PO_NO,")
                    sbCmdText.Append("GR_WH_CODE = @GR_WH_CODE")
                    sbCmdText.Append(" Where VND_CODE=@VND_CODE and STORER_CODE=@STORER_CODE")
                    cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                    cmd.Parameters.AddWithValue("@VND_CODE", row("RO_EDI_PO_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@STORER_CODE", row("STORER_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@GR_TYPE", row("PO_TYPE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@PO_CODE", row("RO_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@GR_DOC_TYPE", row("PO_TYPE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@GR_DOC_NO", row("RO_CODE").ToString.Trim)
                    cmd.Parameters.AddWithValue("@GR_EDI_PO_NO", row("RO_EDI_PO_NO").ToString.Trim)
                    cmd.Parameters.AddWithValue("@GR_WH_CODE", row("RO_WH_CODE").ToString.Trim)
                    cmd.CommandType = System.Data.CommandType.Text
                    cmd.ExecuteScalar()


                    SQLString = "SELECT grd.* FROM WMS_GOODSRCV_D grd Inner Join WMS_GOODSRCV gr ON gr.GR_CODE=grd.GR_CODE and gr.STORER_CODE=grd.STORER_CODE and gr.IMP_CODE=grd.IMP_CODE WHERE grd.GR_CODE='" & row("GR_CODE").ToString.Trim & "'  "
                    Dim codtNPOGRD As New DataTable
                    codtNPOGRD = gDB.getDataTable(SQLString, gConn, transaction)
                    For Each rows As DataRow In codtNPOGRD.Rows
                        Dim txtUpdate As Int16 = 0

                        SQLString = "SELECT * FROM WMS_REPLENISH_D WHERE RO_CODE='" & row("RO_CODE").ToString.Trim & "'"
                        Dim codtRPDNPO As New DataTable
                        codtRPDNPO = gDB.getDataTable(SQLString, gConn, transaction)
                        For Each rowRpd As DataRow In codtRPDNPO.Rows

                            If txtUpdate = 0 Then
                                If rowRpd("ROD_ITM_CODE").ToString.Trim = rows("GRD_ITM_CODE").ToString.Trim Then
                                    UpdateSql = "UPDATE WMS_GOODSRCV_D SET GRD_PALLET_NO ='" & rowRpd("ROD_PALLET_NO").ToString.Trim & "',GRD_REF_NO='" & rowRpd("ROD_SEQ").ToString.Trim & "' WHERE GR_CODE='" & row("GR_CODE").ToString.Trim & "' and GRD_SEQ='" & rowRpd("ROD_SEQ").ToString.Trim & "' and GRD_ITM_CODE='" & rowRpd("ROD_ITM_CODE").ToString.Trim & "'"
                                    gDB.amendData(UpdateSql, gConn, transaction)

                                    UpdateSql = "UPDATE WMS_GOODSRCV_PA SET GRA_PALLET_NO ='" & rowRpd("ROD_PALLET_NO").ToString.Trim & "' WHERE GR_CODE='" & row("GR_CODE").ToString.Trim & "' and GRA_SEQ='" & rowRpd("ROD_SEQ").ToString.Trim & "' and GRA_ITM_CODE='" & rowRpd("ROD_ITM_CODE").ToString.Trim & "'"
                                    gDB.amendData(UpdateSql, gConn, transaction)
                                    txtUpdate = 1
                                End If

                            End If

                        Next

                        If txtUpdate = 0 Then

                            txtUpdate = 1
                        End If

                    Next

                Next

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted PO Header Row :  " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated PO Header Row : " & updateCount.ToString.Trim & "</p>"
                strAlert &= "<p>Inserted PO Details Row : " & insertCountD.ToString.Trim & "</p>"
                strAlert &= "<p>Updated PO Details Row : " & updateCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>PO data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "PO data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import TO
    Public Sub ImportTOData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        'Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        'Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select ACTION,DOC_TYPE,DOC_ID,TRANSACTION_ID,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS = 'NEW' and DOC_TYPE='TO' ORDER BY TRANSACTION_ID ASC"

            Dim codtl As New DataTable
            codtl = gDB.getDataTable(SQLString, gConn, transaction)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim TOInsert As Int16 = 0

                For Each rowD As DataRow In codtl.Rows
                    If (rowD("ACTION").ToString.Trim) = "NEW" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then
                        SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID from EBS_WMS_TRANS_ITX_TO WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                    "WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode Is Nothing Or dtROACTIONCode.Rows.Count <= 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    Dim POCat As String
                                    SQLString = "Select IsNull(STO_CAT,'') AS PO_CAT from WMS_STORER where STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "'"
                                    Dim dtCAT As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtCAT Is Nothing Or dtCAT.Rows.Count <= 0 Then
                                        POCat = ""
                                    Else
                                        POCat = dtCAT.Rows(0)("PO_CAT").ToString
                                    End If

                                    'for PO DATE'
                                    Dim PODATE As String
                                    SQLString = "Select CREATION_DATE from EBS_WMS_TRANS_ITX_ACTION " &
                    "WHERE ACTION <> 'CLOSED' and DOC_TYPE='TO' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString.Trim & "'"
                                    Dim dtPODATE As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtPODATE Is Nothing Or dtPODATE.Rows.Count <= 0 Then
                                        PODATE = ""
                                    Else
                                        PODATE = dtPODATE.Rows(0)("CREATION_DATE").ToString
                                    End If

                                    Dim docId As String
                                    SQLString = "Select IsNULL(DOC_ID,0)DOC_ID from EBS_WMS_TRANS_ITX_ACTION where DOC_TYPE='" & rowD("DOC_TYPE").ToString.Trim & "' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString.Trim & "'"
                                    Dim dtDoc As New DataTable
                                    dtDoc = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtDoc Is Nothing Or dtDoc.Rows.Count <= 0 Then
                                        docId = 0
                                    Else
                                        docId = dtDoc.Rows(0)("DOC_ID").ToString.Trim
                                    End If


                                    'for REPLENISH '
                                    SQLString = "Select * from WMS_REPLENISH " &
                      "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                                        'INSERT
                                        nextNo = DB.getDocNo("RO", gConn, transaction)
                                        Session("TOnxtNo") = nextNo
                                        SQLString = "Insert into WMS_REPLENISH(IMP_CODE,STORER_CODE,RO_CODE,RO_DATE,RO_EDI_PO_NO,RO_SEAL_NO,RO_REF_NO,RO_STATUS,PO_TYPE,RO_TYPE,PO_CAT,RO_WH_CODE,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE, @STORER_CODE,@RO_CODE,@RO_DATE,@RO_EDI_PO_NO,@RO_SEAL_NO,@RO_REF_NO,@RO_STATUS,@PO_TYPE,@RO_TYPE,@PO_CAT,@RO_WH_CODE,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                        cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_CODE", Session("TOnxtNo"))
                                        cmd.Parameters.AddWithValue("@RO_DATE", PODATE)
                                        cmd.Parameters.AddWithValue("@RO_EDI_PO_NO", docId)
                                        cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("TO_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_STATUS", "NEW")
                                        cmd.Parameters.AddWithValue("@PO_TYPE", rowD("DOC_TYPE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_TYPE", rowD("DOC_TYPE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@PO_CAT", POCat)
                                        cmd.Parameters.AddWithValue("@RO_WH_CODE", "FG01")
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        insertCount = insertCount + 1

                                        'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                        TOInsert = 1

                                        'for REPLENISH details'
                                        SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID,LINE_ID,LINE_NUMBER,LOT_NUMBER,ITEM_ID,QUANTITY,UOM_CODE,ITEM_NUMBER from EBS_WMS_TRANS_ITX_TO Where TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "' and TO_HEADER_ID='" & rowH("TO_HEADER_ID").ToString() & "'"
                                        'TO_SUBINVENTORY_CODE
                                        Dim codtTemp As New DataTable
                                        codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                        If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                            For Each rowDtl As DataRow In codtTemp.Rows
                                                'for ITEM NAME'
                                                SQLString = "select Top(1) ITM_NAME ITEM_DESCRIPTION from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                                                Dim dtITM As DataTable
                                                dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                                                Dim ITMName As String
                                                If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                    ITMName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                                Else
                                                    ITMName = ""
                                                End If

                                                Dim RdSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("TOnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                    RdSeq = 0
                                                Else
                                                    RdSeq = dtRPSD.Rows(0)("ROD_SEQ").ToString
                                                End If

                                                'FOR SNO.'
                                                Dim RdDispSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_DISP_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("TOnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                Dim dtDISP As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtDISP Is Nothing Or dtDISP.Rows.Count <= 0 Then
                                                    RdDispSeq = 0
                                                Else
                                                    RdDispSeq = dtDISP.Rows(0)("ROD_DISP_SEQ").ToString
                                                End If

                                                SQLString = "Select * from WMS_REPLENISH_D " &
                             "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("TO_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO = '" & (rowDtl("LINE_ID").ToString.Trim) & "'"
                                                Dim codtlRep As DataTable
                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                    'INSERT
                                                    'Parameterized Parameter
                                                    SQLString = "Insert into WMS_REPLENISH_D(IMP_CODE,ROD_SEQ,ROD_DISP_SEQ,ROD_PALLET_NO,STORER_CODE,ROD_PACK_KEY,RO_CODE,ROD_ITM_CODE,ROD_UOM,ROD_QTY,ROD_DOC_NO,ROD_SERIES_NO,ROD_REF_NO,ROD_STATUS,ROD_BATCH_NO,ROD_SKU_NO,ROD_ITM_NAME,ROD_WH_CODE,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE,@ROD_SEQ,@ROD_DISP_SEQ,@ROD_PALLET_NO,@STORER_CODE,@ROD_PACK_KEY,@RO_CODE,@ROD_ITM_CODE,@ROD_UOM,@ROD_QTY,@ROD_DOC_NO,@ROD_SERIES_NO,@ROD_REF_NO,@ROD_STATUS,@ROD_BATCH_NO,@ROD_SKU_NO,@ROD_ITM_NAME,@ROD_WH_CODE,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                    cmd.Parameters.AddWithValue("@ROD_SEQ", RdSeq)
                                                    cmd.Parameters.AddWithValue("@ROD_DISP_SEQ", RdDispSeq)
                                                    cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("LINE_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PACK_KEY", "1")
                                                    cmd.Parameters.AddWithValue("@RO_CODE", Session("TOnxtNo"))
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("UOM_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("TO_HEADER_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("LINE_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_STATUS", "NEW")
                                                    cmd.Parameters.AddWithValue("@ROD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                                                    'cmd.Parameters.AddWithValue("@ROD_WH_CODE", rowDtl("TO_SUBINVENTORY_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowDtl("ITEM_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_NAME", ITMName)
                                                    cmd.Parameters.AddWithValue("@ROD_WH_CODE", "FG01")
                                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                    cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                                    cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                                    cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                                    cmd.CommandType = System.Data.CommandType.Text
                                                    cmd.ExecuteScalar()
                                                    insertCountD = insertCountD + 1
                                                End If

                                            Next

                                        End If

                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql, gConn, transaction)

                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)
                            lblMSG.Text = "TO data has been already updated"

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then
                        SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID from EBS_WMS_TRANS_ITX_TO WHERE TO_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then

                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                            "WHERE RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CANCEL' WHERE RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("TO_HEADER_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CLOSED" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then
                        SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID from EBS_WMS_TRANS_ITX_TO WHERE TO_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then

                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                        "WHERE RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CLOSED' WHERE RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("TO_HEADER_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If


                        'ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then
                        '    'UPDATE
                        '    SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID from EBS_WMS_TRANS_ITX_TO"
                        '    Dim codt As New DataTable
                        '    Dim adpEBSS As New SqlDataAdapter(SQLString, gConnEBS)
                        '    adpEBSS.Fill(codt)
                        '    For Each rowH As DataRow In codt.Rows
                        '        SQLString = "Select * from WMS_REPLENISH " &
                        '          "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                        '        Dim dtROCode As DataTable
                        '        dtROCode = gDB.getDataTable(SQLString)
                        '        If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                        '            sbCmdText = New StringBuilder()
                        '            sbCmdText.Append("Update WMS_REPLENISH Set ")
                        '            'sbCmdText.Append("RO_STATUS = @RO_STATUS,")
                        '            sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        '            sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '            sbCmdText.Append(" Where RO_SEAL_NO=@RO_SEAL_NO and RO_REF_NO=@RO_REF_NO and STORER_CODE=@STORER_CODE")
                        '            cmd = New SqlCommand(sbCmdText.ToString(), gConn)
                        '            cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        '            cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                        '            cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                        '            cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("TO_HEADER_ID").ToString.Trim)
                        '            'cmd.Parameters.AddWithValue("@RO_STATUS", "NEW")
                        '            cmd.Parameters.AddWithValue("@SYS_LUB", "EBS")
                        '            cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        '            cmd.CommandType = System.Data.CommandType.Text
                        '            cmd.ExecuteScalar()
                        '            updateCount = updateCount + 1
                        '        End If

                        '        SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID,LINE_ID,LOT_NUMBER,LINE_NUMBER,ITEM_ID,QUANTITY,UOM_CODE,ITEM_NUMBER from EBS_WMS_TRANS_ITX_TO Where TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "' and TO_HEADER_ID='" & rowH("TO_HEADER_ID").ToString() & "'"
                        '        Dim codtTemp As New DataTable
                        '        Dim adapter As New SqlDataAdapter(SQLString, gConnEBS)
                        '        adapter.Fill(codtTemp)
                        '        For Each rowDtl As DataRow In codtTemp.Rows
                        '            Dim ItemName As String
                        '            SQLString = "Select IsNUll(ITM_NAME,'')ITM_NAME from WMS_ITEM where ITM_CODE='" & rowDtl("ITEM_ID").ToString.Trim & "' and IMP_CODE='WMS' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                        '            Dim dtITM As DataTable = gDB.getDataTable(SQLString)
                        '            If dtITM Is Nothing Or dtITM.Rows.Count <= 0 Then
                        '                ItemName = ""
                        '            Else
                        '                ItemName = dtITM.Rows(0)("ITM_NAME").ToString.Trim
                        '            End If

                        '            SQLString = "Select * from WMS_REPLENISH_D " &
                        '         "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("TO_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO = '" & (rowDtl("LINE_ID").ToString.Trim) & "'"
                        '            Dim codtlRep As DataTable
                        '            codtlRep = gDB.getDataTable(SQLString)
                        '            If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                        '                sbCmdText = New StringBuilder()
                        '                sbCmdText.Append("Update WMS_REPLENISH_D Set ")
                        '                sbCmdText.Append("ROD_REF_NO = @ROD_REF_NO,")
                        '                sbCmdText.Append("STORER_CODE = @STORER_CODE,")
                        '                sbCmdText.Append("ROD_ITM_CODE = @ROD_ITM_CODE,")
                        '                sbCmdText.Append("ROD_UOM = @ROD_UOM,")
                        '                sbCmdText.Append("ROD_QTY = @ROD_QTY,")
                        '                'sbCmdText.Append("ROD_STATUS = @ROD_STATUS,")
                        '                sbCmdText.Append("ROD_BATCH_NO = @ROD_BATCH_NO,")
                        '                sbCmdText.Append("ROD_SKU_NO = @ROD_SKU_NO,")
                        '                sbCmdText.Append("ROD_ITM_NAME = @ROD_ITM_NAME,")
                        '                sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        '                sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '                sbCmdText.Append(" Where ROD_DOC_NO=@ROD_DOC_NO and ROD_SERIES_NO=@ROD_SERIES_NO and ROD_PALLET_NO=@ROD_PALLET_NO")
                        '                cmd = New SqlCommand(sbCmdText.ToString(), gConn)
                        '                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        '                cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("LINE_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("LINE_NUMBER").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("UOM_CODE").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("TO_HEADER_ID").ToString.Trim)
                        '                'cmd.Parameters.AddWithValue("@ROD_STATUS", "NEW")
                        '                cmd.Parameters.AddWithValue("@ROD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowDtl("ITEM_NUMBER").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@ROD_ITM_NAME", ItemName)
                        '                cmd.Parameters.AddWithValue("@SYS_LUB", "EBS")
                        '                cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        '                cmd.CommandType = System.Data.CommandType.Text
                        '                cmd.ExecuteScalar()
                        '                updateCountD = updateCountD + 1
                        '            End If

                        '        Next
                        '    Next

                        'ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then

                        '    SQLString = "Select TRANSACTION_ID,TO_HEADER_ID,IO_ID from EBS_WMS_TRANS_ITX_TO"
                        '    Dim codt As New DataTable
                        '    Dim adpEBSS As New SqlDataAdapter(SQLString, gConnEBS)
                        '    adpEBSS.Fill(codt)
                        '    For Each rowH As DataRow In codt.Rows
                        '        SQLString = "Select * from WMS_REPLENISH " &
                        '          "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                        '        Dim dtROCode As DataTable
                        '        dtROCode = gDB.getDataTable(SQLString)
                        '        If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                        '            SQLString = "Update WMS_REPLENISH set RO_STATUS='CLOSED'"
                        '            gDB.amendData(SQLString)
                        '        End If
                        '    Next

                    End If

                Next

                'For INTERFACE LOG
                If TOInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "TO_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_TO_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_REPLENISH WHERE PO_TYPE='TO'"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='TO'"
                    Dim TOITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If TOITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            Else
                lblMSG.Text = "TO data has been already updated"

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted TO Header Row : " & insertCount.ToString.Trim & "</p>"
                'strAlert &= "<p>Updated TO Header Row : " & updateCount.ToString.Trim & "</p>"
                strAlert &= "<p>Inserted TO Details Row : " & insertCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>Updated TO Details Row : " & updateCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>TO data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "TO data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import IPR
    Public Sub ImportIPRData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select TRANSACTION_ID,ACTION,DOC_TYPE,DOC_ID,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS <> 'COMPLETED' and STATUS <> 'ERROR' and DOC_TYPE='IPR' ORDER BY TRANSACTION_ID ASC"

            Dim codtl As New DataTable
            codtl = gDB.getDataTable(SQLString, gConn, transaction)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim IPRInsert As Int16 = 0

                For Each rowD As DataRow In codtl.Rows
                    If (rowD("ACTION").ToString.Trim) = "NEW" And (rowD("DOC_TYPE").ToString.Trim) = "IPR" Then
                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,DEST_IO_ID,IPR_NO,SUBINVENTORY_CODE from EBS_WMS_TRANS_ITX_IPR_HEADER WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                            SQLString = "Select * from WMS_REPLENISH " &
                    "WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode Is Nothing Or dtROACTIONCode.Rows.Count <= 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    'For PO_CAT
                                    Dim POCat As String
                                    SQLString = "Select IsNull(STO_CAT,'') AS PO_CAT from WMS_STORER where STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "'"
                                    Dim dtCAT As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtCAT Is Nothing Or dtCAT.Rows.Count <= 0 Then
                                        POCat = ""
                                    Else
                                        POCat = dtCAT.Rows(0)("PO_CAT").ToString
                                    End If

                                    'for PO DATE'
                                    Dim PODATE As String
                                    SQLString = "Select CREATION_DATE from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='IPR' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString.Trim & "'"
                                    Dim dtPODATE As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtPODATE Is Nothing Or dtPODATE.Rows.Count <= 0 Then
                                        PODATE = ""
                                    Else
                                        PODATE = dtPODATE.Rows(0)("CREATION_DATE").ToString
                                    End If

                                    SQLString = "Select * from WMS_REPLENISH " &
                          "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                                        'INSERT
                                        nextNo = DB.getDocNo("RO", gConn, transaction)
                                        Session("IPRnxtNo") = nextNo
                                        SQLString = "Insert into WMS_REPLENISH(IMP_CODE,STORER_CODE,RO_CODE,RO_DATE,RO_EDI_PO_NO,RO_SEAL_NO,RO_REF_NO,RO_STATUS,PO_TYPE,RO_TYPE,PO_CAT,RO_WH_CODE,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE, @STORER_CODE,@RO_CODE,@RO_DATE,@RO_EDI_PO_NO,@RO_SEAL_NO,@RO_REF_NO,@RO_STATUS,@PO_TYPE,@RO_TYPE,@PO_CAT,@RO_WH_CODE,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                        cmd.Parameters.AddWithValue("@STORER_CODE", rowH("DEST_IO_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_CODE", Session("IPRnxtNo"))
                                        cmd.Parameters.AddWithValue("@RO_DATE", PODATE)
                                        cmd.Parameters.AddWithValue("@RO_EDI_PO_NO", rowH("IPR_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("IPR_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_STATUS", "NEW")
                                        cmd.Parameters.AddWithValue("@PO_TYPE", rowD("DOC_TYPE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@RO_TYPE", rowD("DOC_TYPE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@PO_CAT", POCat)
                                        cmd.Parameters.AddWithValue("@RO_WH_CODE", rowH("SUBINVENTORY_CODE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        insertCount = insertCount + 1

                                        'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                        IPRInsert = 1


                                        'FOR IPR Details'
                                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,DEST_IO_ID,IPR_NO,IPR_LINE_ID,ITEM_ID,ITEM_NUMBER,QUANTITY,IPR_LINE_NUMBER,IPR_UOM_CODE from EBS_WMS_TRANS_ITX_IPR_DETAIL Where TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "' and IPR_HEADER_ID='" & rowH("IPR_HEADER_ID").ToString() & "'"
                                        'LOT_NUMBER, TO_IO_ID
                                        Dim codtTemp As New DataTable
                                        codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                        If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                            For Each rowDtl As DataRow In codtTemp.Rows
                                                'FOR SEQ NO.'
                                                Dim RdSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("IPRnxtNo") + "' and IMP_CODE='WMS'"
                                                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                    RdSeq = 0
                                                Else
                                                    RdSeq = dtRPSD.Rows(0)("ROD_SEQ").ToString
                                                End If

                                                'FOR SNO.'
                                                Dim RdDispSeq As String
                                                SQLString = "Select IsNUll(Count(RO_CODE),0)+1 AS ROD_DISP_SEQ from WMS_REPLENISH_D where RO_CODE='" + Session("IPRnxtNo") + "' and IMP_CODE='WMS'"
                                                Dim dtDISP As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                If dtDISP Is Nothing Or dtDISP.Rows.Count <= 0 Then
                                                    RdDispSeq = 0
                                                Else
                                                    RdDispSeq = dtDISP.Rows(0)("ROD_DISP_SEQ").ToString
                                                End If

                                                'for ITEM NAME'
                                                SQLString = "select Top(1) ITM_NAME ITEM_DESCRIPTION,ITM_SKU_NO ITEM_NUMBER,ITM_DRAWING_NO RELATED_ITEM_NUMBER from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("DEST_IO_ID").ToString.Trim & "'"
                                                Dim dtITM As DataTable
                                                dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                                                Dim ItemName As String
                                                If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                    ItemName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                                Else
                                                    ItemName = ""
                                                End If

                                                Dim ItmCode As String
                                                ItmCode = rowDtl("ITEM_ID").ToString.Trim

                                                Dim ITMSKU As String
                                                If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                    If dtITM.Rows(0)("RELATED_ITEM_NUMBER").ToString.Trim IsNot "" Then
                                                        ITMSKU = dtITM.Rows(0)("RELATED_ITEM_NUMBER").ToString.Trim

                                                        SQLString = "select Top(1) ITM_CODE as ITEM_ID from WMS_ITEM where ITM_SKU_NO = '" & ITMSKU & "' and STORER_CODE='" & rowDtl("DEST_IO_ID").ToString.Trim & "'"
                                                        Dim dtITMCode As DataTable
                                                        dtITMCode = gDB.getDataTable(SQLString, gConn, transaction)
                                                        If dtITMCode IsNot Nothing AndAlso dtITMCode.Rows.Count > 0 Then
                                                            ItmCode = dtITMCode.Rows(0)("ITEM_ID").ToString.Trim
                                                        Else
                                                            ItmCode = ""
                                                        End If
                                                    Else
                                                        ITMSKU = dtITM.Rows(0)("ITEM_NUMBER").ToString.Trim
                                                    End If
                                                Else
                                                    ITMSKU = ""
                                                End If


                                                SQLString = "Select * from WMS_REPLENISH_D " &
                                 "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("IPR_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO='" & (rowDtl("IPR_LINE_ID").ToString.Trim) & "' and ROD_REF_NO='" & (rowDtl("IPR_LINE_NUMBER").ToString.Trim) & "'"
                                                Dim codtlRep As DataTable
                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                    'INSERT
                                                    'Parameterized Parameter
                                                    SQLString = "Insert into WMS_REPLENISH_D(IMP_CODE,ROD_SEQ,ROD_DISP_SEQ,STORER_CODE,ROD_PACK_KEY,RO_CODE,ROD_ITM_CODE,ROD_SKU_NO,ROD_UOM,ROD_QTY,ROD_REF_NO,ROD_DOC_NO,ROD_SERIES_NO,ROD_PALLET_NO,ROD_STATUS,ROD_WH_CODE,ROD_ITM_NAME,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE,@ROD_SEQ,@ROD_DISP_SEQ,@STORER_CODE,@ROD_PACK_KEY,@RO_CODE,@ROD_ITM_CODE,@ROD_SKU_NO,@ROD_UOM,@ROD_QTY,@ROD_REF_NO,@ROD_DOC_NO,@ROD_SERIES_NO,@ROD_PALLET_NO,@ROD_STATUS,@ROD_WH_CODE,@ROD_ITM_NAME,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                                    'ROD_BATCH_NO,@ROD_BATCH_NO
                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                    cmd.Parameters.AddWithValue("@ROD_SEQ", RdSeq)
                                                    cmd.Parameters.AddWithValue("@ROD_DISP_SEQ", RdDispSeq)
                                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("DEST_IO_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PACK_KEY", "1")
                                                    cmd.Parameters.AddWithValue("@RO_CODE", Session("IPRnxtNo"))
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_CODE", ItmCode)
                                                    cmd.Parameters.AddWithValue("@ROD_SKU_NO", ITMSKU)
                                                    cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("IPR_UOM_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("IPR_LINE_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("IPR_HEADER_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("IPR_LINE_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_STATUS", "NEW")
                                                    'cmd.Parameters.AddWithValue("@ROD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_WH_CODE", "FG01")
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_NAME", ItemName)
                                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                    cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                                    cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                                                    cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                                    cmd.CommandType = System.Data.CommandType.Text
                                                    cmd.ExecuteScalar()
                                                    insertCountD = insertCountD + 1
                                                End If

                                            Next

                                        End If

                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql, gConn, transaction)

                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)
                            lblMSG.Text = "IPR data has been already updated"

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "IPR" Then
                        'UPDATE
                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,DEST_IO_ID,IPR_NO from EBS_WMS_TRANS_ITX_IPR_HEADER WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                            SQLString = "Select * from WMS_REPLENISH " &
                        "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then
                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                          "WHERE STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        sbCmdText = New StringBuilder()
                                        sbCmdText.Append("Update WMS_REPLENISH Set ")
                                        sbCmdText.Append("RO_EDI_PO_NO = @RO_EDI_PO_NO,")
                                        sbCmdText.Append("PO_TYPE = @PO_TYPE,")
                                        sbCmdText.Append("RO_TYPE = @RO_TYPE,")
                                        sbCmdText.Append("RO_WH_CODE = @RO_WH_CODE,")
                                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                                        sbCmdText.Append(" Where RO_SEAL_NO=@RO_SEAL_NO and RO_REF_NO=@RO_REF_NO and STORER_CODE=@STORER_CODE")
                                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                        cmd.Parameters.AddWithValue("@STORER_CODE", rowH("DEST_IO_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_SEAL_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_REF_NO", rowH("IPR_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RO_EDI_PO_NO", rowH("IPR_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@PO_TYPE", rowD("DOC_TYPE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@RO_TYPE", rowD("DOC_TYPE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@RO_WH_CODE", rowH("SUBINVENTORY_CODE").ToString.Trim())
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        updateCount = updateCount + 1


                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        'UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        'gDB.amendData(UpdateSql, gConn, transaction)


                                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,IPR_NO,IPR_LINE_ID,ITEM_ID,ITEM_NUMBER,QUANTITY,IPR_LINE_NUMBER,IPR_UOM_CODE from EBS_WMS_TRANS_ITX_IPR_DETAIL Where TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "' and IPR_HEADER_ID='" & rowH("IPR_HEADER_ID").ToString() & "'"
                                        Dim codtTemp As New DataTable
                                        codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                        If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                            For Each rowDtl As DataRow In codtTemp.Rows
                                                SQLString = "Select * from WMS_REPLENISH_D " &
                             "WHERE ROD_DOC_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowDtl("IPR_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO='" & (rowDtl("IPR_LINE_ID").ToString.Trim) & "'and ROD_REF_NO='" & (rowDtl("IPR_LINE_NUMBER").ToString.Trim) & "'"
                                                Dim codtlRep As DataTable
                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                                                    sbCmdText = New StringBuilder()
                                                    sbCmdText.Append("Update WMS_REPLENISH_D Set ")
                                                    sbCmdText.Append("STORER_CODE = @STORER_CODE,")
                                                    sbCmdText.Append("ROD_PACK_KEY = @ROD_PACK_KEY,")
                                                    sbCmdText.Append("ROD_ITM_CODE = @ROD_ITM_CODE,")
                                                    'sbCmdText.Append("ROD_SKU_NO = @ROD_SKU_NO,")
                                                    sbCmdText.Append("ROD_UOM = @ROD_UOM,")
                                                    sbCmdText.Append("ROD_QTY = @ROD_QTY,")
                                                    'sbCmdText.Append("ROD_BATCH_NO = @ROD_BATCH_NO,")
                                                    sbCmdText.Append("ROD_WH_CODE = @ROD_WH_CODE,")
                                                    sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                                                    sbCmdText.Append("SYS_LUD = @SYS_LUD")
                                                    sbCmdText.Append(" Where ROD_DOC_NO=@ROD_DOC_NO and ROD_SERIES_NO=@ROD_SERIES_NO and ROD_PALLET_NO=@ROD_PALLET_NO and ROD_REF_NO=@ROD_REF_NO")
                                                    cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("DEST_IO_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PACK_KEY", "1")
                                                    cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                                                    'cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowDtl("ITEM_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_UOM", rowDtl("IPR_UOM_CODE").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_QTY", rowDtl("QUANTITY").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_REF_NO", rowDtl("IPR_LINE_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowDtl("IPR_HEADER_ID").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowDtl("IPR_LINE_ID").ToString.Trim)
                                                    'cmd.Parameters.AddWithValue("@ROD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                                                    cmd.Parameters.AddWithValue("@ROD_WH_CODE", "FG01")
                                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                    cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                                    cmd.CommandType = System.Data.CommandType.Text
                                                    cmd.ExecuteScalar()
                                                    updateCountD = updateCountD + 1
                                                End If

                                            Next

                                        End If

                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql, gConn, transaction)

                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "IPR" Then
                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,IO_ID,IPR_NO from EBS_WMS_TRANS_ITX_IPR_HEADER WHERE IPR_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then

                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                                  "WHERE RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CANCEL' WHERE RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                        gDB.amendData(UpdateSql, gConn, transaction)

                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Cancel must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CANCEL'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CLOSED" And (rowD("DOC_TYPE").ToString.Trim) = "IPR" Then
                        SQLString = "Select TRANSACTION_ID,IPR_HEADER_ID,IO_ID,IPR_NO from EBS_WMS_TRANS_ITX_IPR_HEADER WHERE IPR_HEADER_ID='" & (rowD("DOC_ID").ToString.Trim) & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH " &
                                "WHERE RO_REF_NO ='" & (rowD("DOC_ID").ToString.Trim) & "'"
                            '"WHERE RO_SEAL_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then

                                For Each rowH As DataRow In codt.Rows
                                    SQLString = "Select * from WMS_REPLENISH " &
                                  "WHERE RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                    'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and
                                    Dim dtROCode As DataTable
                                    dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                    If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                        SQLString = "Update WMS_REPLENISH set RO_STATUS='CLOSED' WHERE RO_REF_NO='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "'"
                                        'RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("DEST_IO_ID").ToString.Trim) & "' and 
                                        gDB.amendData(SQLString, gConn, transaction)

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowH("IPR_HEADER_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If
                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                                gDB.amendData(UpdateSql, gConn, transaction)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Closed must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE DOC_ID='" & (rowD("DOC_ID").ToString.Trim) & "' AND ACTION='CLOSED'"
                            gDB.amendData(UpdateSql, gConn, transaction)

                        End If

                    End If

                Next

                'For INTERFACE LOG
                If IPRInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "IPR_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_IPR_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_REPLENISH WHERE PO_TYPE='IPR'"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='IPR'"
                    Dim IPRITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If IPRITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            Else
                lblMSG.Text = "IPR data has been already updated"
            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted IPR Header Row : " & insertCount.ToString.Trim & "</p>"
                'strAlert &= "<p>Updated IPR Header Row : " & updateCount.ToString.Trim & "</p>"
                strAlert &= "<p>Inserted IPR Details Row : " & insertCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>Updated IPR Details Row : " & updateCountD.ToString.Trim & "</p>"
                'strAlert &= "<p>IPR data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "IPR data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import STOCK RETURN
    Public Sub ImportSRData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        'lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select TRANSACTION_ID,ACTION,DOC_TYPE,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS = 'NEW' and DOC_TYPE='SO_RETURN'"

            Dim codtl As New DataTable
            codtl = gDB.getDataTable(SQLString, gConn, transaction)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim SRInsert As Int16 = 0

                For Each rowD As DataRow In codtl.Rows
                    If (rowD("ACTION").ToString.Trim) = "NEW" And (rowD("DOC_TYPE").ToString.Trim) = "SO_RETURN" Then
                        SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN WHERE TRANSACTION_ID='" & rowD("TRANSACTION_ID").ToString.Trim & "'"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)
                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codt IsNot Nothing AndAlso codt.Rows.Count > 0) Then

                            For Each rowH As DataRow In codt.Rows
                                'SQLString = "Select * from WMS_STOCK_RETURN WHERE RT_STATUS!='CANCELLED' and RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "'"
                                SQLString = "Select * from WMS_STOCK_RETURN SR inner join WMS_STOCK_RETURN_D SRD on SR.RT_CODE=SRD.RT_CODE where RT_STATUS!='CANCELLED' and RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "' and  RTD_ITM_CODE = '" & rowH("ITEM_ID").ToString.Trim & "' "
                                Dim codtSR As DataTable
                                codtSR = gDB.getDataTable(SQLString, gConn, transaction)

                                SQLString = "Select * from WMS_STOCK_RETURN SR inner join WMS_STOCK_RETURN_D SRD on SR.RT_CODE=SRD.RT_CODE where RT_STATUS!='CANCELLED' and RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "' and  RTD_ITM_CODE = '" & rowH("ITEM_ID").ToString.Trim & "' and RT_CONS_CODE = '" & rowH("SO_RETURN_NO").ToString.Trim & "' "
                                Dim codtRetCode As DataTable
                                codtRetCode = gDB.getDataTable(SQLString, gConn, transaction)

                                If codtSR Is Nothing Or codtSR.Rows.Count <= 0 Or codtRetCode IsNot Nothing Or codtRetCode.Rows.Count > 0 Then
                                    'INSERT
                                    'FOR WMS_STOCK_RETURN'
                                    nextNo = DB.getDocNo("SR", gConn, transaction)
                                    Session("SRnxtNo") = nextNo
                                    SQLString = "Insert into WMS_STOCK_RETURN(IMP_CODE,STORER_CODE,RT_CODE,RT_TYPE,RT_STATUS,RT_DATE,RT_CUS_CODE,RT_CONS_CODE,RT_WH,RT_REF_NO,RT_REF_NO2,RT_REF_DOC_NO,RT_APPROVE_CODE,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE, @STORER_CODE,@RT_CODE,@RT_TYPE,@RT_STATUS,@RT_DATE,@RT_CUS_CODE,@RT_CONS_CODE,@RT_WH,@RT_REF_NO,@RT_REF_NO2,@RT_REF_DOC_NO,@RT_APPROVE_CODE,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_CODE", Session("SRnxtNo"))
                                    cmd.Parameters.AddWithValue("@RT_TYPE", "DORETURN")
                                    cmd.Parameters.AddWithValue("@RT_STATUS", "NEW")
                                    cmd.Parameters.AddWithValue("@RT_DATE", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_CUS_CODE", rowH("CUSTOMER_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_CONS_CODE", rowH("SO_RETURN_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_WH", rowH("SUBINVENTORY_CODE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_REF_NO2", rowH("LINE_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_REF_DOC_NO", rowH("SO_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RT_APPROVE_CODE", rowH("SO_HEADER_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                    cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                    cmd.Parameters.AddWithValue("@SYS_CD", rowH("CREATION_DATE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                    insertCount = insertCount + 1

                                    'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                    UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                    gDB.amendData(UpdateSql, gConn, transaction)
                                    SRInsert = 1

                                    'FOR WMS_STOCK_RETURN_D'
                                    SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN Where SO_HEADER_ID='" & rowH("SO_HEADER_ID").ToString() & "' and SO_NO='" & rowH("SO_NO").ToString() & "' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "'"
                                    Dim codtTemp As New DataTable
                                    codtTemp = gDB.getDataTable(SQLString, gConn, transaction)

                                    If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                        For Each rowDtl As DataRow In codtTemp.Rows
                                            Dim RdSeq As String
                                            SQLString = "Select IsNUll(Count(RT_CODE),0)+1 AS RTD_SEQ from WMS_STOCK_RETURN_D where RT_CODE='" + Session("SRnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                            Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                            If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                RdSeq = 0
                                            Else
                                                RdSeq = dtRPSD.Rows(0)("RTD_SEQ").ToString
                                            End If

                                            'for ITEM NAME'
                                            SQLString = "select Top(1) ITM_NAME as ITEM_DESCRIPTION from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                                            Dim dtITM As DataTable
                                            dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                                            Dim ITMName As String
                                            If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                ITMName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                            Else
                                                ITMName = ""
                                            End If


                                            SQLString = "Select * from WMS_STOCK_RETURN_D " &
                             "WHERE RTD_REF_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE = '" & (rowDtl("IO_ID").ToString.Trim) & "' and RTD_PALLET_NO='" & (rowDtl("LINE_ID").ToString.Trim) & "' and SAP_MAT_DOC_ITEM='" & rowDtl("SO_HEADER_ID").ToString.Trim & "' and SAP_MAT_DOC_NO='" & rowDtl("SO_NO").ToString.Trim & "'"
                                            Dim codtlRep As DataTable
                                            codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                            If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                'INSERT
                                                'Parameterized Parameter
                                                SQLString = "Insert into WMS_STOCK_RETURN_D(IMP_CODE,STORER_CODE,RTD_SEQ,RTD_PACK_KEY,RT_CODE,RTD_STATUS,RTD_REF_NO,RTD_PALLET_NO,SAP_MAT_DOC_ITEM,SAP_MAT_DOC_NO,RTD_ITM_CODE,RTD_ITM_NAME,RTD_LOC_WH,RTD_DRUM_LV,RTD_KG,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD,RTD_WH,RTD_LOC) VALUES (@IMP_CODE, @STORER_CODE,@RTD_SEQ,@RTD_PACK_KEY,@RT_CODE,@RTD_STATUS,@RTD_REF_NO,@RTD_PALLET_NO,@SAP_MAT_DOC_ITEM,@SAP_MAT_DOC_NO,@RTD_ITM_CODE,@RTD_ITM_NAME,@RTD_LOC_WH,@RTD_DRUM_LV,@RTD_KG,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD,@RTD_WH,@RTD_LOC)"
                                                cmd = New SqlCommand(SQLString, gConn, transaction)
                                                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_SEQ", RdSeq)
                                                cmd.Parameters.AddWithValue("@RTD_PACK_KEY", "1")
                                                cmd.Parameters.AddWithValue("@RT_CODE", Session("SRnxtNo"))
                                                cmd.Parameters.AddWithValue("@RTD_STATUS", "NEW")
                                                cmd.Parameters.AddWithValue("@RTD_REF_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_PALLET_NO", rowDtl("LINE_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SAP_MAT_DOC_ITEM", rowDtl("SO_HEADER_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SAP_MAT_DOC_NO", rowDtl("SO_NO").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_ITM_NAME", ITMName)
                                                cmd.Parameters.AddWithValue("@RTD_LOC_WH", rowDtl("SUBINVENTORY_CODE").ToString.Trim)
                                                'cmd.Parameters.AddWithValue("@RTD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                                                'cmd.Parameters.AddWithValue("@RTD_UOM2", rowDtl("UOM_CODE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_DRUM_LV", rowDtl("UOM_CODE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_KG", rowDtl("QUANTITY").ToString.Trim)
                                                'cmd.Parameters.AddWithValue("@RTD_RCV_QTY", rowDtl("QUANTITY").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                                cmd.Parameters.AddWithValue("@SYS_CD", rowH("CREATION_DATE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_WH", "QCFG")
                                                cmd.Parameters.AddWithValue("@RTD_LOC", "0A0101FG")
                                                cmd.CommandType = System.Data.CommandType.Text
                                                cmd.ExecuteScalar()
                                                insertCountD = insertCountD + 1
                                            End If

                                        Next

                                    End If

                                ElseIf (codtSR IsNot Nothing AndAlso codtSR.Rows.Count > 0 AndAlso codtSR.Rows(0)("RT_STATUS").ToString.ToUpper = "APPROVED" Or codtSR.Rows(0)("RT_STATUS").ToString.ToUpper = "POSTED") Then
                                    SQLString = "Select top(1) * from WMS_STOCK_RETURN SR inner join WMS_STOCK_RETURN_D SRD on SR.RT_CODE=SRD.RT_CODE where RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "' and  RTD_ITM_CODE = '" & rowH("ITEM_ID").ToString.Trim & "' and ISNULL(RT_CONS_CODE,'')='' and (isnull(SRD.RTD_PALLET_NO,'')='' or isnull(SRD.RTD_PALLET_NO,'')='000') and " & rowH("QUANTITY").ToString.Trim & " >= isnull(SRD.RTD_RCV_QTY,0) order by SRD.RTD_RCV_QTY desc "
                                    Dim codtSReturn As New DataTable
                                    codtSReturn = gDB.getDataTable(SQLString, gConn, transaction)
                                    If codtSReturn IsNot Nothing AndAlso codtSReturn.Rows.Count > 0 Then
                                        sbCmdText = New StringBuilder()
                                        sbCmdText.Append("Update WMS_STOCK_RETURN Set ")
                                        sbCmdText.Append("RT_DATE = @RT_DATE,")
                                        sbCmdText.Append("RT_CUS_CODE = @RT_CUS_CODE,")
                                        sbCmdText.Append("RT_CONS_CODE = @RT_CONS_CODE,")
                                        sbCmdText.Append("RT_WH = @RT_WH,")
                                        sbCmdText.Append("RT_REF_NO = @RT_REF_NO,")
                                        sbCmdText.Append("RT_REF_NO2 = @RT_REF_NO2,")
                                        sbCmdText.Append("RT_APPROVE_CODE = @RT_APPROVE_CODE,")
                                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                                        sbCmdText.Append(" Where RT_REF_DOC_NO=@RT_REF_DOC_NO")
                                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                        cmd.Parameters.AddWithValue("@RT_DATE", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_CUS_CODE", rowH("CUSTOMER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_CONS_CODE", rowH("SO_RETURN_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_WH", rowH("SUBINVENTORY_CODE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_REF_NO2", rowH("LINE_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_REF_DOC_NO", rowH("SO_NO").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@RT_APPROVE_CODE", rowH("SO_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        updateCount = updateCount + 1

                                        For Each rowDtlSReturn As DataRow In codtSReturn.Rows
                                            SQLString = "Select * from WMS_STOCK_RETURN_D where RT_CODE='" & rowDtlSReturn("RT_CODE").ToString.Trim & "' and RTD_ITM_CODE = '" & (rowH("ITEM_ID").ToString.Trim) & "'"
                                            Dim codtlRep As DataTable
                                            codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                            If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                                                'UPDATE
                                                sbCmdText = New StringBuilder()
                                                sbCmdText.Append("Update WMS_STOCK_RETURN_D Set ")
                                                sbCmdText.Append("STORER_CODE = @STORER_CODE,")
                                                sbCmdText.Append("RTD_REF_NO = @RTD_REF_NO,")
                                                sbCmdText.Append("RTD_PALLET_NO = @RTD_PALLET_NO,")
                                                sbCmdText.Append("SAP_MAT_DOC_ITEM = @SAP_MAT_DOC_ITEM,")
                                                sbCmdText.Append("SAP_MAT_DOC_NO = @SAP_MAT_DOC_NO,")
                                                sbCmdText.Append("RTD_LOC_WH = @RTD_LOC_WH,")
                                                sbCmdText.Append("RTD_KG = @RTD_KG,")
                                                sbCmdText.Append("RTD_DRUM_LV = @RTD_DRUM_LV,")
                                                sbCmdText.Append("SYS_LUD = @SYS_LUD")
                                                sbCmdText.Append(" Where RTD_ITM_CODE=@RTD_ITM_CODE and RT_CODE = @RT_CODE ")
                                                cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                                cmd.Parameters.AddWithValue("@RT_CODE", rowDtlSReturn("RT_CODE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_PALLET_NO", rowH("LINE_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SAP_MAT_DOC_ITEM", rowH("SO_HEADER_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SAP_MAT_DOC_NO", rowH("SO_NO").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_ITM_CODE", rowH("ITEM_ID").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_LOC_WH", rowH("SUBINVENTORY_CODE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_KG", rowH("QUANTITY").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@RTD_DRUM_LV", rowH("UOM_CODE").ToString.Trim)
                                                cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                                cmd.CommandType = System.Data.CommandType.Text
                                                cmd.ExecuteScalar()

                                            End If

                                            updateCountD = updateCountD + 1
                                        Next

                                        'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql, gConn, transaction)
                                    End If

                                Else
                                    uiFun.displayMsg(Me, "", "SO Number found but status Not Approved Or Posted", Session("gLang"))
                                End If

                            Next

                            'Else
                            '    UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            '    gDB.amendData(UpdateSql, gConn, transaction)
                            'End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql, gConn, transaction)
                            'lblMSG.Text = "Stock Return data has been already updated"
                        End If

                        'ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "SO_RETURN" Then
                        '    SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN WHERE TRANSACTION_ID='" & rowD("TRANSACTION_ID").ToString.Trim & "'"
                        '    Dim codt As New DataTable
                        '    Dim adpEBSS As New SqlDataAdapter(SQLString, gConnEBS)
                        '    adpEBSS.Fill(codt)
                        '    For Each rowH As DataRow In codt.Rows
                        '        SQLString = "Select * from WMS_STOCK_RETURN WHERE RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "'"
                        '        Dim codtSR As DataTable
                        '        codtSR = gDB.getDataTable(SQLString)
                        '        If codtSR IsNot Nothing AndAlso (codtSR.Rows(0)("RT_STATUS") = "NEW") Then
                        '            SQLString = "Select * from WMS_STOCK_RETURN WHERE RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "'"
                        '            Dim codtSReturn As New DataTable
                        '            codtSReturn = gDB.getDataTable(SQLString)
                        '            If codtSReturn IsNot Nothing AndAlso codtSReturn.Rows.Count > 0 Then
                        '                sbCmdText = New StringBuilder()
                        '                sbCmdText.Append("Update WMS_STOCK_RETURN Set ")
                        '                sbCmdText.Append("RT_DATE = @RT_DATE,")
                        '                sbCmdText.Append("RT_CUS_CODE = @RT_CUS_CODE,")
                        '                sbCmdText.Append("RT_CONS_CODE = @RT_CONS_CODE,")
                        '                sbCmdText.Append("RT_WH = @RT_WH,")
                        '                sbCmdText.Append("RT_REF_NO = @RT_REF_NO,")
                        '                sbCmdText.Append("RT_REF_NO2 = @RT_REF_NO2,")
                        '                sbCmdText.Append("RT_APPROVE_CODE = @RT_APPROVE_CODE,")
                        '                sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '                sbCmdText.Append(" Where RT_REF_DOC_NO=@RT_REF_DOC_NO")
                        '                cmd = New SqlCommand(sbCmdText.ToString(), gConn)
                        '                cmd.Parameters.AddWithValue("@RT_DATE", rowH("LAST_UPDATE_DATE").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_CUS_CODE", rowH("CUSTOMER_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_CONS_CODE", rowH("REF_CUSTOMER_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_WH", rowH("SUBINVENTORY_CODE").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_REF_NO2", rowH("LINE_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_REF_DOC_NO", rowH("SO_NO").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@RT_APPROVE_CODE", rowH("SO_HEADER_ID").ToString.Trim)
                        '                cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                        '                cmd.CommandType = System.Data.CommandType.Text
                        '                cmd.ExecuteScalar()
                        '                updateCount = updateCount + 1
                        '            End If

                        '            SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN Where SO_HEADER_ID='" & rowH("SO_HEADER_ID").ToString() & "' and SO_NO='" & rowH("SO_NO").ToString() & "' and TRANSACTION_ID='" & rowH("TRANSACTION_ID").ToString() & "'"
                        '            Dim codtTemp As New DataTable
                        '            Dim adapter As New SqlDataAdapter(SQLString, gConnEBS)
                        '            adapter.Fill(codtTemp)
                        '            For Each rowDtl As DataRow In codtTemp.Rows
                        '                SQLString = "Select * from WMS_STOCK_RETURN_D " &
                        '         "WHERE RTD_REF_NO='" & (rowDtl("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE = '" & (rowDtl("IO_ID").ToString.Trim) & "' and RTD_PALLET_NO='" & (rowDtl("LINE_ID").ToString.Trim) & "' and SAP_MAT_DOC_ITEM='" & rowDtl("SO_HEADER_ID").ToString.Trim & "' and SAP_MAT_DOC_NO='" & rowDtl("SO_NO").ToString.Trim & "'"
                        '                Dim codtlRep As DataTable
                        '                codtlRep = gDB.getDataTable(SQLString)
                        '                If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                        '                    'UPDATE
                        '                    sbCmdText = New StringBuilder()
                        '                    sbCmdText.Append("Update WMS_STOCK_RETURN_D Set ")
                        '                    sbCmdText.Append("RTD_ITM_CODE = @RTD_ITM_CODE,")
                        '                    sbCmdText.Append("RTD_LOC_WH = @RTD_LOC_WH,")
                        '                    sbCmdText.Append("RTD_BATCH_NO = @RTD_BATCH_NO,")
                        '                    sbCmdText.Append("RTD_UOM2 = @RTD_UOM2,")
                        '                    sbCmdText.Append("RTD_RCV_QTY = @RTD_RCV_QTY,")
                        '                    sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        '                    sbCmdText.Append(" Where RTD_REF_NO=@RTD_REF_NO and STORER_CODE=@STORER_CODE and RTD_PALLET_NO=@RTD_PALLET_NO and SAP_MAT_DOC_ITEM=@SAP_MAT_DOC_ITEM and SAP_MAT_DOC_NO=@SAP_MAT_DOC_NO")
                        '                    cmd = New SqlCommand(sbCmdText.ToString(), gConn)
                        '                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_REF_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_PALLET_NO", rowDtl("LINE_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@SAP_MAT_DOC_ITEM", rowDtl("SO_HEADER_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@SAP_MAT_DOC_NO", rowDtl("SO_NO").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_LOC_WH", rowDtl("SUBINVENTORY_CODE").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_BATCH_NO", rowDtl("LOT_NUMBER").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_UOM2", rowDtl("UOM_CODE").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@RTD_RCV_QTY", rowDtl("QUANTITY").ToString.Trim)
                        '                    cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                        '                    cmd.CommandType = System.Data.CommandType.Text
                        '                    cmd.ExecuteScalar()
                        '                    updateCountD = updateCountD + 1
                        '                End If

                        '            Next

                        '        End If

                        '    Next

                        'ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "SO_RETURN" Then
                        '    SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN WHERE TRANSACTION_ID='" & rowD("TRANSACTION_ID").ToString.Trim & "'"
                        '    Dim codt As New DataTable
                        '    Dim adpEBSS As New SqlDataAdapter(SQLString, gConnEBS)
                        '    adpEBSS.Fill(codt)
                        '    For Each rowH As DataRow In codt.Rows
                        '        SQLString = "Select * from WMS_STOCK_RETURN WHERE RT_REF_DOC_NO='" & rowH("SO_NO").ToString.Trim & "'"
                        '        Dim codtSR As DataTable
                        '        codtSR = gDB.getDataTable(SQLString)
                        '        If codtSR IsNot Nothing AndAlso codtSR.Rows.Count > 0 Then
                        '            SQLString = "Update WMS_STOCK_RETURN set RT_STATUS='CLOSED'"
                        '            gDB.amendData(SQLString)
                        '        Else

                        '        End If
                        '    Next

                    End If

                Next

                'For INTERFACE LOG
                If SRInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "SR_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_SR_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_STOCK_RETURN"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                            "WHERE ACTION <> 'CLOSED' and DOC_TYPE='SO_RETURN'"
                    Dim SRITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If SRITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            Else
                'lblMSG.Text = "Stock Return data has been already updated"
            End If

            transaction.Commit()



        Catch ex As Exception
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import SO
    Public Sub ImportSOData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As New DataTable
        Dim nextNo As String
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now
            CurrTrans = ""
            SQLString = "Select TRANSACTION_ID,ACTION,DOC_TYPE,STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE STATUS = 'NEW' and DOC_TYPE='SO'"

            codtl = gDB.getDataTable(SQLString)

            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='INPROGRESS',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE STATUS = 'NEW' and DOC_TYPE='SO' "
            gDB.amendData(UpdateSql)

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                Dim SOInsert As Int16 = 0

                For Each rowD As DataRow In codtl.Rows
                    CurrTrans = "Action: " & rowD("TRANSACTION_ID").ToString.Trim
                    If (rowD("ACTION").ToString.Trim) = "NEW" And (rowD("DOC_TYPE").ToString.Trim) = "SO" Then
                        SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_HEADER where TRANSACTION_ID='" & rowD("TRANSACTION_ID").ToString.Trim & "' "
                        Dim codtlTemp As New DataTable
                        codtlTemp = gDB.getDataTable(SQLString)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codtlTemp IsNot Nothing AndAlso codtlTemp.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_CUST_ORDER " &
                    "WHERE CO_CUS_REF_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode Is Nothing Or dtROACTIONCode.Rows.Count <= 0 Then
                                For Each rowH As DataRow In codtlTemp.Rows
                                    'for EBS_WMS_TRANS_ITX_SO_DETAIL(CREDIT_HOLD)
                                    SQLString = "Select Top(1) CREDIT_HOLD from EBS_WMS_TRANS_ITX_SO_DETAIL WHERE Transaction_ID='" & rowH("TRANSACTION_ID") & "' "
                                    Dim codtTempp As New DataTable
                                    codtTempp = gDB.getDataTable(SQLString)

                                    If codtTempp IsNot Nothing AndAlso codtTempp.Rows.Count > 0 Then
                                        'for WMS_CUST_ORDER'
                                        SQLString = "Select * from WMS_CUST_ORDER " &
                    "WHERE CO_PROJECT_NO='" & (rowH("SO_HEADER_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and CO_CUS_REF_NO='" & rowH("TRANSACTION_ID").ToString.Trim & "'"
                                        Dim codtlCO As New DataTable
                                        codtlCO = gDB.getDataTable(SQLString, gConn, transaction)
                                        If codtlCO Is Nothing Or codtlCO.Rows.Count <= 0 Then
                                            'INSERT
                                            'Parameterized Parameter
                                            nextNo = DB.getDocNo("CO", gConn, transaction)
                                            Session("COnxtNo") = nextNo
                                            SQLString = "Insert into WMS_CUST_ORDER(IMP_CODE,STORER_CODE,CO_CODE,CUS_NAME,CO_DATE,CUS_CODE,CO_CUS_REF_NO,CO_PROJECT_NO,CO_STATUS,CO_ADDR1,CO_ADDR2,CO_ADDR3,CO_AREA_DEL,CO_FTRACK_NO,CO_INV_NO,ROUTE_ID,CO_SENDER_REGION,SYS_LUB,SYS_LUD,SYS_CD,SYS_CB,CO_EDI_SIR_NO,CO_SENDER) VALUES (@IMP_CODE,@STORER_CODE,@CO_CODE,@CUS_NAME,@CO_DATE,@CUS_CODE,@CO_CUS_REF_NO,@CO_PROJECT_NO,@CO_STATUS,@CO_ADDR1,@CO_ADDR2,@CO_ADDR3,@CO_AREA_DEL,@CO_FTRACK_NO,@CO_INV_NO,@ROUTE_ID,@CO_SENDER_REGION,@SYS_LUB,@SYS_LUD,@SYS_CD,@SYS_CB,@CO_EDI_SIR_NO,@CO_SENDER)"
                                            cmd = New SqlCommand(SQLString, gConn, transaction)
                                            cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                            cmd.Parameters.AddWithValue("@STORER_CODE", rowH("IO_ID").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_CODE", Session("COnxtNo"))
                                            cmd.Parameters.AddWithValue("@CO_DATE", rowH("REQUEST_DATE").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CUS_CODE", rowH("ACCOUNT_NUMBER").ToString.Trim)
                                            'for CUSTOMER NAME'
                                            SQLString = "select Top(1) Customer_NAME from EBS_WMS_CUSTOMER_MASTER where ACCOUNT_NUMBER = '" & rowH("ACCOUNT_NUMBER").ToString.Trim & "'"
                                            Dim dtCusname As DataTable
                                            dtCusname = gDB.getDataTable(SQLString, gConn, transaction)
                                            Dim CustomerName As String
                                            If dtCusname IsNot Nothing AndAlso dtCusname.Rows.Count > 0 Then
                                                CustomerName = dtCusname.Rows(0)("Customer_NAME").ToString.Trim
                                            Else
                                                CustomerName = ""
                                            End If
                                            cmd.Parameters.AddWithValue("@CUS_NAME", CustomerName)

                                            If (codtTempp.Rows(0)("CREDIT_HOLD").ToString.Trim) = "True" Then
                                                cmd.Parameters.AddWithValue("@CO_SENDER_REGION", "True")
                                            ElseIf (codtTempp.Rows(0)("CREDIT_HOLD").ToString.Trim) = "False" Then
                                                cmd.Parameters.AddWithValue("@CO_SENDER_REGION", "False")
                                            Else
                                                cmd.Parameters.AddWithValue("@CO_SENDER_REGION", "False")
                                            End If

                                            cmd.Parameters.AddWithValue("@CO_CUS_REF_NO", rowH("TRANSACTION_ID").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_PROJECT_NO", rowH("SO_HEADER_ID").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_STATUS", "NEW")
                                            cmd.Parameters.AddWithValue("@CO_ADDR1", rowH("ADDRESS1").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_ADDR2", rowH("ADDRESS2").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_ADDR3", rowH("ADDRESS1").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_AREA_DEL", rowH("ADDRESS4").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_FTRACK_NO", rowH("TRIAL_TYPE").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_INV_NO", rowH("SO_NO").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@ROUTE_ID", rowH("ROUTE").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                            cmd.Parameters.AddWithValue("@SYS_LUD", rowH("LAST_UPDATE_DATE").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@SYS_CD", rowH("CREATION_DATE").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                                            cmd.Parameters.AddWithValue("@CO_EDI_SIR_NO", rowH("BATCH_NO").ToString.Trim)
                                            cmd.Parameters.AddWithValue("@CO_SENDER", rowH("EBS_USER").ToString.Trim)
                                            cmd.CommandType = System.Data.CommandType.Text
                                            cmd.ExecuteScalar()
                                            insertCount = insertCount + 1

                                            'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='UPLOADED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowH("TRANSACTION_ID").ToString.Trim) & "'"
                                            gDB.amendData(UpdateSql)
                                            SOInsert = 1

                                            'for WMS_CUST_ORDER_D'
                                            SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_DETAIL WHERE Transaction_ID='" & rowH("TRANSACTION_ID") & "' "
                                            Dim codtTemp As New DataTable
                                            codtTemp = gDB.getDataTable(SQLString)

                                            If codtTemp IsNot Nothing AndAlso codtTemp.Rows.Count > 0 Then

                                                For Each rowDtl As DataRow In codtTemp.Rows

                                                    Dim CdLotNo As String
                                                    SQLString = "Select LOT_NUMBER from EBS_WMS_TRANS_ITX_SO_DETAIL_LOT where TRANSACTION_ID='" & rowDtl("TRANSACTION_ID") & "' and SO_HEADER_ID='" & rowDtl("SO_HEADER_ID").ToString.Trim & "' and SO_LINE_ID='" & (rowDtl("SO_LINE_ID").ToString.Trim) & "'"
                                                    Dim codtSOLOT As New DataTable
                                                    codtSOLOT = gDB.getDataTable(SQLString)

                                                    If codtSOLOT Is Nothing Or codtSOLOT.Rows.Count <= 0 Then
                                                        CdLotNo = ""
                                                    Else
                                                        CdLotNo = codtSOLOT.Rows(0)("LOT_NUMBER").ToString.Trim
                                                    End If

                                                    'for ITEM NAME & ITEM NUMBER'
                                                    SQLString = "select Top(1) ITM_NAME as ITEM_DESCRIPTION,ITM_CODE as ITEM_NUMBER from WMS_ITEM where ITM_CODE = '" & rowDtl("ITEM_ID").ToString.Trim & "' and STORER_CODE='" & rowDtl("IO_ID").ToString.Trim & "'"
                                                    Dim dtITM As DataTable
                                                    dtITM = gDB.getDataTable(SQLString, gConn, transaction)
                                                    Dim ITMName, ITEMNUMBER As String
                                                    If dtITM IsNot Nothing AndAlso dtITM.Rows.Count > 0 Then
                                                        ITMName = dtITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                                        ITEMNUMBER = dtITM.Rows(0)("ITEM_NUMBER").ToString.Trim
                                                    Else
                                                        ITMName = ""
                                                        ITEMNUMBER = ""
                                                    End If

                                                    'For last lot no
                                                    SQLString = "select Top(1) LOT_NUMBER from WMS_LAST_LOTS where CUSTOMER_CODE = '" & rowH("ACCOUNT_NUMBER").ToString.Trim & "' and ITEM='" & ITEMNUMBER & "'"
                                                    Dim dtLastLotNo As DataTable
                                                    dtLastLotNo = gDB.getDataTable(SQLString, gConn, transaction)

                                                    'get warehouse'
                                                    Dim WarehouseType As String
                                                    SQLString = "Select SUBINVENTORY_CODE from EBS_WMS_TRANS_ITX_SO_DETAIL_LOT where TRANSACTION_ID='" & rowDtl("TRANSACTION_ID") & "' and SO_HEADER_ID='" & rowDtl("SO_HEADER_ID").ToString.Trim & "' and SO_LINE_ID='" & (rowDtl("SO_LINE_ID").ToString.Trim) & "'"
                                                    Dim codtSOLOTWH As New DataTable
                                                    codtSOLOTWH = gDB.getDataTable(SQLString)

                                                    If codtSOLOTWH Is Nothing Or codtSOLOTWH.Rows.Count <= 0 Then
                                                        WarehouseType = ""
                                                    Else
                                                        WarehouseType = codtSOLOTWH.Rows(0)("SUBINVENTORY_CODE").ToString
                                                    End If

                                                    If WarehouseType = "FGLOT" Or WarehouseType = "DIRECTLOT" Then
                                                        SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_DETAIL_LOT WHERE Transaction_ID='" & rowDtl("TRANSACTION_ID") & "' "
                                                        Dim codtTempLot As New DataTable
                                                        codtTempLot = gDB.getDataTable(SQLString)

                                                        If codtTempLot IsNot Nothing AndAlso codtTempLot.Rows.Count > 0 Then
                                                            For Each rowLot As DataRow In codtTempLot.Rows
                                                                Dim CdSeq As String
                                                                SQLString = "Select IsNUll(Count(CO_CODE),0)+1 AS COD_SEQ from WMS_CUST_ORDER_D where CO_CODE='" + Session("COnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                                Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                                If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                                    CdSeq = 0
                                                                Else
                                                                    CdSeq = dtRPSD.Rows(0)("COD_SEQ").ToString
                                                                End If

                                                                'FOR SNO.'
                                                                Dim RdDispSeq As String
                                                                SQLString = "Select IsNUll(Count(CO_CODE),0)+1 AS COD_DISP_SEQ from WMS_CUST_ORDER_D where CO_CODE='" + Session("COnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                                Dim dtDISP As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                                If dtDISP Is Nothing Or dtDISP.Rows.Count <= 0 Then
                                                                    RdDispSeq = 0
                                                                Else
                                                                    RdDispSeq = dtDISP.Rows(0)("COD_DISP_SEQ").ToString
                                                                End If

                                                                'for ITEM NAME'
                                                                SQLString = "select Top(1) ITM_NAME as ITEM_DESCRIPTION from WMS_ITEM where ITM_CODE = '" & rowLot("ITEM_ID").ToString.Trim & "'"
                                                                Dim dtLotITM As DataTable
                                                                dtLotITM = gDB.getDataTable(SQLString, gConn, transaction)
                                                                Dim LotITMName As String
                                                                If dtLotITM IsNot Nothing AndAlso dtLotITM.Rows.Count > 0 Then
                                                                    LotITMName = dtLotITM.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                                                                Else
                                                                    LotITMName = ""
                                                                End If

                                                                SQLString = "Select * from WMS_CUST_ORDER_D " &
                                       "WHERE COD_JOB_NO='" & rowLot("TRANSACTION_ID").ToString.Trim & "' and COD_REF_NO='" & rowLot("SO_HEADER_ID").ToString.Trim & "' and COD_PACKING='" & rowLot("SO_LINE_ID").ToString.Trim & "'"
                                                                Dim codtlRep As DataTable
                                                                codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                                If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                                    'INSERT
                                                                    'Parameterized Parameter
                                                                    SQLString = "Insert into WMS_CUST_ORDER_D(IMP_CODE,STORER_CODE,COD_SEQ,COD_DISP_SEQ,CO_CODE,COD_PACK_KEY,COD_JOB_NO,COD_REF_NO,COD_ITM_CODE,COD_ITM_DESC,COD_PACKING,COD_QTY,COD_UOM,COD_WH_CODE,COD_BATCH_NO,COD_PALLET_NO,COD_PLANT,COD_CARTON_NO,COD_TICKET_NO) VALUES (@IMP_CODE, @STORER_CODE,@COD_SEQ,@COD_DISP_SEQ,@CO_CODE,@COD_PACK_KEY,@COD_JOB_NO,@COD_REF_NO,@COD_ITM_CODE,@COD_ITM_DESC,@COD_PACKING,@COD_QTY,@COD_UOM,@COD_WH_CODE,@COD_BATCH_NO,@COD_PALLET_NO,@COD_PLANT,@COD_CARTON_NO,@COD_TICKET_NO)"
                                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                                    cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                                    cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_SEQ", CdSeq)
                                                                    cmd.Parameters.AddWithValue("@COD_DISP_SEQ", RdDispSeq)
                                                                    cmd.Parameters.AddWithValue("@CO_CODE", Session("COnxtNo"))
                                                                    cmd.Parameters.AddWithValue("@COD_PACK_KEY", "1")
                                                                    cmd.Parameters.AddWithValue("@COD_JOB_NO", rowLot("TRANSACTION_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_REF_NO", rowLot("SO_HEADER_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_ITM_CODE", rowLot("ITEM_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_ITM_DESC", LotITMName)
                                                                    cmd.Parameters.AddWithValue("@COD_PACKING", rowLot("SO_LINE_ID").ToString.Trim)
                                                                    Dim LastLotNo As String = getLastLot(rowLot("ITEM_ID").ToString.Trim, rowDtl("IO_ID").ToString.Trim, rowH("ACCOUNT_NUMBER").ToString.Trim)
                                                                    If LastLotNo <> "" Then
                                                                        cmd.Parameters.AddWithValue("@COD_PALLET_NO", LastLotNo)
                                                                    ElseIf dtLastLotNo IsNot Nothing AndAlso dtLastLotNo.Rows.Count > 0 Then
                                                                        cmd.Parameters.AddWithValue("@COD_PALLET_NO", dtLastLotNo.Rows(0)("LOT_NUMBER").ToString.Trim)
                                                                    Else
                                                                        cmd.Parameters.AddWithValue("@COD_PALLET_NO", "000")
                                                                    End If
                                                                    cmd.Parameters.AddWithValue("@COD_QTY", rowLot("LOT_PRIMARY_QUANTITY").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_UOM", rowLot("ORDERED_PRIMARY_UOM").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_WH_CODE", rowLot("SUBINVENTORY_CODE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_BATCH_NO", CdLotNo)

                                                                    cmd.Parameters.AddWithValue("@COD_PLANT", rowDtl("SO_LINE_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_CARTON_NO", rowDtl("SALESMAN").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@COD_TICKET_NO", rowDtl("CUST_PO_NUMBER").ToString.Trim)

                                                                    cmd.CommandType = System.Data.CommandType.Text
                                                                    cmd.ExecuteScalar()
                                                                    insertCountD = insertCountD + 1
                                                                End If

                                                            Next

                                                        End If

                                                    Else
                                                        Dim CdSeq As String
                                                        SQLString = "Select IsNUll(Count(CO_CODE),0)+1 AS COD_SEQ from WMS_CUST_ORDER_D where CO_CODE='" + Session("COnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                        Dim dtRPSD As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                        If dtRPSD Is Nothing Or dtRPSD.Rows.Count <= 0 Then
                                                            CdSeq = 0
                                                        Else
                                                            CdSeq = dtRPSD.Rows(0)("COD_SEQ").ToString
                                                        End If

                                                        'FOR SNO.'
                                                        Dim RdDispSeq As String
                                                        SQLString = "Select IsNUll(Count(CO_CODE),0)+1 AS COD_DISP_SEQ from WMS_CUST_ORDER_D where CO_CODE='" + Session("COnxtNo") + "' and IMP_CODE='WMS' and STORER_CODE='" & (rowDtl("IO_ID").ToString.Trim) & "'"
                                                        Dim dtDISP As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                        If dtDISP Is Nothing Or dtDISP.Rows.Count <= 0 Then
                                                            RdDispSeq = 0
                                                        Else
                                                            RdDispSeq = dtDISP.Rows(0)("COD_DISP_SEQ").ToString
                                                        End If

                                                        SQLString = "Select * from WMS_CUST_ORDER_D " &
                               "WHERE COD_JOB_NO='" & rowDtl("TRANSACTION_ID").ToString.Trim & "' and COD_REF_NO='" & rowDtl("SO_HEADER_ID").ToString.Trim & "' and COD_PACKING='" & rowDtl("SO_LINE_ID").ToString.Trim & "'"
                                                        Dim codtlRep As DataTable
                                                        codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                                        If codtlRep Is Nothing Or codtlRep.Rows.Count <= 0 Then
                                                            'INSERT
                                                            'Parameterized Parameter
                                                            SQLString = "Insert into WMS_CUST_ORDER_D(IMP_CODE,STORER_CODE,COD_SEQ,COD_DISP_SEQ,CO_CODE,COD_PACK_KEY,COD_JOB_NO,COD_REF_NO,COD_ITM_CODE,COD_ITM_DESC,COD_PACKING,COD_QTY,COD_UOM,COD_BATCH_NO,COD_WH_CODE,COD_PALLET_NO,COD_PLANT,COD_CARTON_NO,COD_TICKET_NO) VALUES (@IMP_CODE, @STORER_CODE,@COD_SEQ,@COD_DISP_SEQ,@CO_CODE,@COD_PACK_KEY,@COD_JOB_NO,@COD_REF_NO,@COD_ITM_CODE,@COD_ITM_DESC,@COD_PACKING,@COD_QTY,@COD_UOM,@COD_BATCH_NO,@COD_WH_CODE,@COD_PALLET_NO,@COD_PLANT,@COD_CARTON_NO,@COD_TICKET_NO)"
                                                            'COD_WH_CODE,@COD_WH_CODE
                                                            cmd = New SqlCommand(SQLString, gConn, transaction)
                                                            cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                                            cmd.Parameters.AddWithValue("@STORER_CODE", rowDtl("IO_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_SEQ", CdSeq)
                                                            cmd.Parameters.AddWithValue("@COD_DISP_SEQ", RdDispSeq)
                                                            cmd.Parameters.AddWithValue("@CO_CODE", Session("COnxtNo"))
                                                            cmd.Parameters.AddWithValue("@COD_PACK_KEY", "1")
                                                            cmd.Parameters.AddWithValue("@COD_JOB_NO", rowDtl("TRANSACTION_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_REF_NO", rowDtl("SO_HEADER_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_ITM_CODE", rowDtl("ITEM_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_ITM_DESC", ITMName)
                                                            cmd.Parameters.AddWithValue("@COD_PACKING", rowDtl("SO_LINE_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_QTY", rowDtl("ORDERED_PRIMARY_QUANTITY").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_UOM", rowDtl("ORDERED_PRIMARY_UOM").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_WH_CODE", rowDtl("SUBINVENTORY_CODE").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_BATCH_NO", CdLotNo)

                                                            cmd.Parameters.AddWithValue("@COD_PLANT", rowDtl("SO_LINE_NUMBER").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_CARTON_NO", rowDtl("SALESMAN").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@COD_TICKET_NO", rowDtl("CUST_PO_NUMBER").ToString.Trim)

                                                            Dim LastLotNo As String = getLastLot(rowDtl("ITEM_ID").ToString.Trim, rowDtl("IO_ID").ToString.Trim, rowH("ACCOUNT_NUMBER").ToString.Trim)
                                                            If LastLotNo <> "" Then
                                                                cmd.Parameters.AddWithValue("@COD_PALLET_NO", LastLotNo)
                                                            ElseIf dtLastLotNo IsNot Nothing AndAlso dtLastLotNo.Rows.Count > 0 Then
                                                                cmd.Parameters.AddWithValue("@COD_PALLET_NO", dtLastLotNo.Rows(0)("LOT_NUMBER").ToString.Trim)
                                                            Else
                                                                cmd.Parameters.AddWithValue("@COD_PALLET_NO", "000")
                                                            End If
                                                            cmd.CommandType = System.Data.CommandType.Text
                                                            cmd.ExecuteScalar()
                                                            insertCountD = insertCountD + 1
                                                        End If

                                                    End If

                                                Next

                                            End If

                                        End If
                                    Else
                                        UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='SO Details not exists',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                        gDB.amendData(UpdateSql)
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='New must not exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql)
                            lblMSG.Text = "SO data has been already updated"

                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "SO" Then
                        SQLString = "Select * from EBS_WMS_TRANS_ITX_TO"
                        Dim codtlTempTO As New DataTable
                        codtlTempTO = gDB.getDataTable(SQLString, gConn, transaction)

                        If (rowD("STATUS").ToString.Trim = "NEW" AndAlso codtlTempTO IsNot Nothing AndAlso codtlTempTO.Rows.Count > 0) Then
                            SQLString = "Select * from WMS_REPLENISH_D " &
                        "WHERE ROD_DOC_NO ='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            Dim dtROACTIONCode As DataTable
                            dtROACTIONCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtROACTIONCode IsNot Nothing AndAlso dtROACTIONCode.Rows.Count > 0 Then

                                For Each rowTO As DataRow In codtlTempTO.Rows
                                    SQLString = "Select * from WMS_REPLENISH_D " &
                             "WHERE ROD_DOC_NO='" & (rowTO("TRANSACTION_ID").ToString.Trim) & "' and ROD_SERIES_NO = '" & (rowTO("TO_HEADER_ID").ToString.Trim) & "' and ROD_PALLET_NO = '" & (rowTO("LINE_ID").ToString.Trim) & "'"
                                    Dim codtlRep As DataTable
                                    codtlRep = gDB.getDataTable(SQLString, gConn, transaction)
                                    If codtlRep IsNot Nothing AndAlso codtlRep.Rows.Count > 0 Then
                                        sbCmdText = New StringBuilder()
                                        sbCmdText.Append("Update WMS_REPLENISH_D Set ")
                                        sbCmdText.Append("ROD_REF_NO = @ROD_REF_NO,")
                                        sbCmdText.Append("STORER_CODE = @STORER_CODE,")
                                        sbCmdText.Append("ROD_ITM_CODE = @ROD_ITM_CODE,")
                                        sbCmdText.Append("ROD_UOM = @ROD_UOM,")
                                        sbCmdText.Append("ROD_QTY = @ROD_QTY,")
                                        sbCmdText.Append("ROD_BATCH_NO = @ROD_BATCH_NO,")
                                        sbCmdText.Append("ROD_SKU_NO = @ROD_SKU_NO,")
                                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                                        sbCmdText.Append(" Where ROD_DOC_NO=@ROD_DOC_NO and ROD_SERIES_NO=@ROD_SERIES_NO and ROD_PALLET_NO=@ROD_PALLET_NO")
                                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                                        cmd.Parameters.AddWithValue("@ROD_PALLET_NO", rowTO("LINE_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_REF_NO", rowTO("LINE_NUMBER").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@STORER_CODE", rowTO("IO_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_ITM_CODE", rowTO("ITEM_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_UOM", rowTO("UOM_CODE").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_QTY", rowTO("QUANTITY").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_DOC_NO", rowTO("TRANSACTION_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_SERIES_NO", rowTO("TO_HEADER_ID").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_BATCH_NO", rowTO("LOT_NUMBER").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@ROD_SKU_NO", rowTO("ITEM_NUMBER").ToString.Trim)
                                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                                        cmd.CommandType = System.Data.CommandType.Text
                                        cmd.ExecuteScalar()
                                        updateCountD = updateCountD + 1
                                    End If

                                Next

                            Else
                                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                                gDB.amendData(UpdateSql)
                            End If

                        Else
                            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("TRANSACTION_ID").ToString.Trim) & "'"
                            gDB.amendData(UpdateSql)
                        End If

                    ElseIf (rowD("ACTION").ToString.Trim) = "CANCEL" And (rowD("DOC_TYPE").ToString.Trim) = "SO" Then
                        SQLString = "Select * from EBS_WMS_TRANS_ITX_TO_DETAIL"
                        Dim codt As New DataTable
                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                        If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                            For Each rowH As DataRow In codt.Rows
                                SQLString = "Select * from WMS_REPLENISH " &
                              "WHERE RO_SEAL_NO ='" & (rowH("TRANSACTION_ID").ToString.Trim) & "' and STORER_CODE='" & (rowH("IO_ID").ToString.Trim) & "' and RO_REF_NO='" & (rowH("TO_HEADER_ID").ToString.Trim) & "'"
                                Dim dtROCode As DataTable
                                dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
                                If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                    SQLString = "Update WMS_REPLENISH set RO_STATUS='CLOSED'"
                                    gDB.amendData(SQLString, gConn, transaction)
                                End If

                            Next

                        End If

                    End If

                Next
                CurrTrans = "SO Download Completed"
                SQLString = "SELECT distinct(isnull(CO_EDI_SIR_NO,'')) SOBatchNo from WMS_CUST_ORDER where CO_FTRACK_NO='TRIAL' and CO_STATUS='NEW'"
                Dim dtBatchList As DataTable
                dtBatchList = gDB.getDataTable(SQLString, gConn, transaction)

                For Each rowBatchNo As DataRow In dtBatchList.Rows
                    CurrTrans = "Batch No: " & rowBatchNo("SOBatchNo").ToString.Trim
                    SQLString = "SELECT 1 As IsTrail from WMS_CUST_ORDER where CO_FTRACK_NO='TRIAL' and CO_STATUS='NEW' and isnull(CO_EDI_SIR_NO,'')='" & rowBatchNo("SOBatchNo").ToString.Trim & "' "
                    Dim dtTotalTrailRecords As DataTable
                    dtTotalTrailRecords = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtTotalTrailRecords IsNot Nothing AndAlso dtTotalTrailRecords.Rows.Count > 0 Then
                        'For Lot Allocation'
                        SQLString = "select * from WMS_CUST_ORDER WHERE CO_STATUS='NEW' and isnull(CO_EDI_SIR_NO,'')='" & rowBatchNo("SOBatchNo").ToString.Trim & "'"
                        Dim dtTotalCORecords As DataTable
                        Dim ebs_user As String = ""
                        Dim storer_code As String = "104"
                        dtTotalCORecords = gDB.getDataTable(SQLString, gConn, transaction)
                        If dtTotalCORecords IsNot Nothing AndAlso dtTotalCORecords.Rows.Count > 0 Then
                            LotAllocationCO(transaction, gConn, rowBatchNo("SOBatchNo").ToString.Trim)
                            ebs_user = dtTotalCORecords.Rows(0)("CO_SENDER").ToString.Trim
                            storer_code = dtTotalCORecords.Rows(0)("STORER_CODE").ToString.Trim
                        End If

                        'Send Allocation Email'
                        'Dim a As Byte
                        'For a = 1 To 3
                        '    If a = 1 Then
                        '        SendEmail(gConn, transaction, "104", "362", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        '    ElseIf a = 2 Then
                        '        SendEmail(gConn, transaction, "105", "0", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        '    ElseIf a = 3 Then
                        '        SendEmail(gConn, transaction, "106", "0", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        '    End If
                        'Next

                        If storer_code = "105" Then
                            SendEmail(gConn, transaction, "105", "0", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        ElseIf storer_code = "106" Then
                            SendEmail(gConn, transaction, "106", "0", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        Else
                            SendEmail(gConn, transaction, "104", "362", rowBatchNo("SOBatchNo").ToString.Trim, ebs_user)
                        End If

                        If rowBatchNo("STORER_TYPE").ToString.Trim = "NEW" Then
                            SQLString = "SELECT  DO_CODE FROM WMS_DELV_ORDER Where ROUTE_ID in ('All','0')  and DO_STATUS!='POSTED' "
                            Dim dtDORecords As DataTable
                            dtDORecords = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtDORecords IsNot Nothing AndAlso dtDORecords.Rows.Count > 0 Then
                                For Each rowDO As DataRow In dtDORecords.Rows
                                    SQLString = "DELETE FROM WMS_DELV_ORDER where DO_CODE='" & rowDO("DO_CODE").ToString() & "'; DELETE FROM WMS_DELV_ORDER_D where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_DO_PICKLIST_D where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_DELV_ORDER_TEMP where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_WAVEPICK_RSVD where DO_CODE='" & rowDO("DO_CODE").ToString() & "'"
                                    gDB.amendData(SQLString, gConn, transaction)
                                Next
                            End If
                        Else
                            SQLString = "SELECT do.DO_CODE FROM WMS_DELV_ORDER do inner join WMS_DELV_ORDER_D dod on do.DO_CODE=dod.DO_CODE inner join WMS_CUST_ORDER co on co.CO_CODE IN (select items from dbo.split(do.DO_CO_CODE,',')) where co.CO_FTRACK_NO='TRIAL' and isnull(co.CO_EDI_SIR_NO,'')='" & rowBatchNo("SOBatchNo").ToString.Trim & "' and co.CO_STATUS!='NEW'  and do.DO_STATUS!='POSTED' "
                            Dim dtDORecords As DataTable
                            dtDORecords = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtDORecords IsNot Nothing AndAlso dtDORecords.Rows.Count > 0 Then
                                For Each rowDO As DataRow In dtDORecords.Rows
                                    SQLString = "DELETE FROM WMS_DELV_ORDER where DO_CODE='" & rowDO("DO_CODE").ToString() & "'; DELETE FROM WMS_DELV_ORDER_D where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_DO_PICKLIST_D where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_DELV_ORDER_TEMP where DO_CODE='" & rowDO("DO_CODE").ToString() & "';DELETE FROM WMS_WAVEPICK_RSVD where DO_CODE='" & rowDO("DO_CODE").ToString() & "'"
                                    gDB.amendData(SQLString, gConn, transaction)
                                Next
                            End If
                        End If

                        'Delete Trail CO '
                        SQLString = "SELECT co.CO_CODE,cod.CO_CODE FROM WMS_CUST_ORDER co INNER JOIN WMS_CUST_ORDER_D cod on co.CO_CODE=cod.CO_CODE where co.CO_FTRACK_NO='TRIAL' and isnull(co.CO_EDI_SIR_NO,'')='" & rowBatchNo("SOBatchNo").ToString.Trim & "'"
                        Dim dtCORecords As DataTable
                        dtCORecords = gDB.getDataTable(SQLString, gConn, transaction)
                        If dtCORecords IsNot Nothing AndAlso dtCORecords.Rows.Count > 0 Then
                            For Each rowCO As DataRow In dtCORecords.Rows
                                SQLString = "DELETE FROM WMS_CUST_ORDER where CO_CODE='" & rowCO("CO_CODE").ToString() & "'; DELETE FROM WMS_CUST_ORDER_D where CO_CODE='" & rowCO("CO_CODE").ToString() & "'"
                                gDB.amendData(SQLString, gConn, transaction)
                            Next
                        End If
                    End If

                Next
                CurrTrans = "Batch Completed"
                'For deleting blank DO master records
                SQLString = "delete from WMS_DELV_ORDER where (select count(1) from WMS_DELV_ORDER_D where WMS_DELV_ORDER_D.DO_CODE=WMS_DELV_ORDER.DO_CODE)=0"
                gDB.amendData(SQLString, gConn, transaction)

                'For INTERFACE LOG
                If SOInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "SO_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_SO_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_CUST_ORDER"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from EBS_WMS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='SO'"
                    Dim SOITXACTION As DataTable = gDB.getDataTable(SQLString)

                    Dim ITFSTATUS As String
                    If SOITXACTION.Rows(0)("STATUS").ToString.Trim = "COMPLETED" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='" & codtl.Rows(0)("DOC_TYPE").ToString.Trim & "' and ITF_TYPE='I'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", (codtl.Rows(0)("DOC_TYPE").ToString.Trim), TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If



            Else
                lblMSG.Text = "SO data has been already updated"
            End If

            transaction.Commit()

            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE STATUS IN ('UPLOADED') and DOC_TYPE='SO' "
            gDB.amendData(UpdateSql)

            'If lblMSG.Text = "" Then
            Dim strAlert As String = "<p>SO Header Row Inserted : " & insertCount.ToString.Trim & "</p>"
            'strAlert &= "<p>Header Row Updated : " & updateCount.ToString.Trim & "</p>"
            strAlert &= "<p>SO Detail Row Inserted : " & insertCountD.ToString.Trim & "</p>"
            'strAlert &= "<p>Detail Row Updated : " & updateCountD.ToString.Trim & "</p>"
            'strAlert &= "<p>SO data has been successfully imported!!</p>"
            lblMSG.Text = strAlert
            'Else
            '    lblMSG.Text = "SO data has been already updated"
            'End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message & "Error Action: " & CurrTrans)
            uiFun.displayMsg(Me, "", ex.Message & "Error Action: " & CurrTrans, Session("gLang"))

            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='NEW',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE STATUS IN ('UPLOADED','INPROGRESS') and DOC_TYPE='SO' "
            gDB.amendData(UpdateSql)

            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    Public Sub ImportINV_TRANSFERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select * FROM WMS_STOCK_TRANSFER WHERE TR_STATUS='POSTED' and EBS_UPDATE_STATUS='N'"
            Dim codtGRCV As New DataTable
            codtGRCV = gDB.getDataTable(SQLString, gConn, transaction)

            If codtGRCV IsNot Nothing AndAlso codtGRCV.Rows.Count > 0 Then
                Dim STKRELOCATIONInsert As Int16 = 0

                For Each rowGRCV As DataRow In codtGRCV.Rows
                    'FOR ACTION BATCH_NO'
                    Dim dtACTIONBatchNo As New DataTable
                    SQLString = "Select REPLACE('WMSINV'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'0000000000000'),0))+1),'WMSINV1','WMSINV') as BATCH_NO From WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER"
                    dtACTIONBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                    'for TRANSACTION_ID as use seq_no'
                    Dim dtSeq As New DataTable
                    SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                    dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select top(1) * from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE DOC_TYPE='INV_TRANSFER'"
                    Dim codtl As New DataTable
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)

                    If 1 = 1 Then 'codtl Is Nothing Or codtl.Rows.Count <= 0
                        'INSERT
                        SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,STATUS,BATCH_NO,CREATION_DATE,LAST_UPDATE_DATE) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@STATUS,@BATCH_NO,@CREATION_DATE,@LAST_UPDATE_DATE)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                        cmd.Parameters.AddWithValue("@ACTION", "NEW")
                        cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@IO_ID", rowGRCV("STORER_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@COMPANY_ID", "83")
                        cmd.Parameters.AddWithValue("@DOC_TYPE", "INV_TRANSFER")
                        cmd.Parameters.AddWithValue("@STATUS", "NEW")
                        cmd.Parameters.AddWithValue("@BATCH_NO", dtACTIONBatchNo.Rows(0)("BATCH_NO").ToString.Trim())
                        cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()

                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_EBS_TRANS_ITX_ACTION Set ")
                        sbCmdText.Append("ACTION = @ACTION,")
                        sbCmdText.Append("IO_ID = @IO_ID,")
                        sbCmdText.Append("COMPANY_ID = @COMPANY_ID,")
                        sbCmdText.Append("DOC_TYPE = @DOC_TYPE,")
                        sbCmdText.Append("BATCH_NO = @BATCH_NO,")
                        sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                        sbCmdText.Append(" Where TRANSACTION_ID = @TRANSACTION_ID")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@ACTION", "WMS")
                        cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@IO_ID", rowGRCV("STORER_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@COMPANY_ID", "83")
                        cmd.Parameters.AddWithValue("@DOC_TYPE", "INV_TRANSFER")
                        cmd.Parameters.AddWithValue("@BATCH_NO", dtACTIONBatchNo.Rows(0)("BATCH_NO").ToString.Trim())
                        cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                    End If

                    'FOR BATCH_NO'
                    Dim dtBatchNo As New DataTable
                    SQLString = "Select REPLACE('WMSINV'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'0000000000000'),0))+1),'WMSINV1','WMSINV') as BATCH_NO From WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER"
                    dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select * FROM WMS_STOCK_TRANSFER_D WHERE TR_CODE='" & rowGRCV("TR_CODE").ToString.Trim & "'"
                    Dim codtGRCV_D As New DataTable
                    codtGRCV_D = gDB.getDataTable(SQLString, gConn, transaction)

                    For Each rowGRCV_D As DataRow In codtGRCV_D.Rows

                        'for seq_no'
                        Dim dtTO As New DataTable
                        SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER"
                        dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                        'for transaction_Id'
                        Dim dtTransaction_Id As New DataTable
                        SQLString = "Select MAX(TRANSACTION_ID) as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION WHERE DOC_TYPE='INV_TRANSFER'"
                        dtTransaction_Id = gDB.getDataTable(SQLString, gConn, transaction)

                        'ITM_SKU_NO'
                        Dim ITM_SKU_NO, ITM_UOM As String
                        Dim dtITMCODE As New DataTable
                        SQLString = "Select ITM_SKU_NO,ITM_UOM from WMS_ITEM WHERE ITM_CODE='" & rowGRCV_D("ITM_CODE").ToString.Trim & "'"
                        dtITMCODE = gDB.getDataTable(SQLString, gConn, transaction)
                        If dtITMCODE IsNot Nothing AndAlso dtITMCODE.Rows.Count > 0 Then
                            ITM_SKU_NO = dtITMCODE.Rows(0)("ITM_SKU_NO").ToString.Trim
                            ITM_UOM = dtITMCODE.Rows(0)("ITM_UOM").ToString.Trim
                        Else
                            ITM_SKU_NO = ""
                            ITM_UOM = ""
                        End If

                        'For STORER 
                        Dim IO_CODE As String
                        SQLString = "select Top(1) IO_CODE from EBS_WMS_COMPANY_MASTER where IO_ID = '" & rowGRCV_D("STORER_CODE").ToString.Trim & "' "
                        Dim dtIOCode As DataTable
                        dtIOCode = gDB.getDataTable(SQLString, gConn, transaction)
                        If dtIOCode IsNot Nothing AndAlso dtIOCode.Rows.Count > 0 Then
                            IO_CODE = dtIOCode.Rows(0)("IO_CODE").ToString.Trim
                        Else
                            IO_CODE = ""
                        End If

                        'for FROM_LOCATOR'
                        Dim FROM_LOCATOR, FRLOCATOR As String
                        Dim TRD_LOC_FR As String = rowGRCV_D("TRD_LOC_FR").ToString.Trim
                        FRLOCATOR = Right(TRD_LOC_FR, 8)
                        FROM_LOCATOR = FRLOCATOR.Substring(0, 2) + "." + FRLOCATOR.Substring(2, 2) + "." + FRLOCATOR.Substring(4, 2) + "." + FRLOCATOR.Substring(6, 2)

                        'for TO_LOCATOR'
                        Dim TO_LOCATOR, TOLOCATOR As String
                        Dim TRD_LOC_TO As String = rowGRCV_D("TRD_LOC_TO").ToString.Trim
                        TOLOCATOR = Right(TRD_LOC_TO, 8)
                        TO_LOCATOR = TOLOCATOR.Substring(0, 2) + "." + TOLOCATOR.Substring(2, 2) + "." + TOLOCATOR.Substring(4, 2) + "." + TOLOCATOR.Substring(6, 2)

                        'INSERT
                        SQLString = "Insert into WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,IO_CODE,FROM_SUBINVENTORY_CODE,FROM_LOCATOR,TO_SUBINVENTORY_CODE,TO_LOCATOR,LOT_NUMBER,ITEM_NUMBER,UOM_CODE,QUANTITY,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@IO_CODE,@FROM_SUBINVENTORY_CODE,@FROM_LOCATOR,@TO_SUBINVENTORY_CODE,@TO_LOCATOR,@LOT_NUMBER,@ITEM_NUMBER,@UOM_CODE,@QUANTITY,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                        cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                        cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtTransaction_Id.Rows(0)("TRANSACTION_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@TRANSACTION_DATE", rowGRCV_D("SYS_CD").ToString.Trim)
                        cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                        cmd.Parameters.AddWithValue("@FROM_SUBINVENTORY_CODE", rowGRCV_D("TRD_LOC_FR").ToString.Trim.Remove(rowGRCV_D("TRD_LOC_FR").ToString.Length - 8, 8))
                        cmd.Parameters.AddWithValue("@FROM_LOCATOR", FROM_LOCATOR)
                        cmd.Parameters.AddWithValue("@TO_SUBINVENTORY_CODE", rowGRCV_D("TRD_LOC_TO").ToString.Trim.Remove(rowGRCV_D("TRD_LOC_TO").ToString.Length - 8, 8))
                        cmd.Parameters.AddWithValue("@TO_LOCATOR", TO_LOCATOR)
                        If rowGRCV_D("TRD_BATCH_NO_FR").ToString.Trim <> "" Then
                            cmd.Parameters.AddWithValue("@LOT_NUMBER", rowGRCV_D("TRD_BATCH_NO_FR").ToString.Trim)
                        Else
                            cmd.Parameters.AddWithValue("@LOT_NUMBER", DBNull.Value)
                        End If
                        cmd.Parameters.AddWithValue("@ITEM_NUMBER", ITM_SKU_NO)
                        cmd.Parameters.AddWithValue("@UOM_CODE", ITM_UOM)
                        cmd.Parameters.AddWithValue("@QUANTITY", rowGRCV_D("TRD_QTY").ToString.Trim)
                        cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", rowGRCV_D("SYS_CD").ToString.Trim)
                        cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@CREATION_DATE", rowGRCV_D("SYS_CD").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1
                        STKRELOCATIONInsert = 1
                    Next

                    SQLString = "UPDATE WMS_STOCK_TRANSFER SET EBS_UPDATE_STATUS ='Y' FROM WMS_STOCK_TRANSFER WHERE TR_CODE='" & rowGRCV("TR_CODE").ToString.Trim & "' and TR_STATUS ='POSTED'"
                    gDB.amendData(SQLString, gConn, transaction)

                Next

                'For INTERFACE LOG
                If STKRELOCATIONInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "STOCK_RELOCATION_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_STOCK_RELOCATION_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='INV_TRANSFER'"
                    Dim SOITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If SOITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='INV_TRANSFER' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "INV_TRANSFER", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            Dim strAlert As String = "<p>INVENTORY TRANSFER Inserted Row : " & insertCount.ToString.Trim & "</p>"
            lblMSG.Text = strAlert

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import PO_GRN
    Public Sub ImportPO_GRNData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select a.* FROM WMS_GOODSRCV a Where a.GR_STATUS ='POSTED' and a.EBS_UPDATE_STATUS='N' and a.GR_DOC_TYPE='PO'"
            Dim codtGRCV As New DataTable
            codtGRCV = gDB.getDataTable(SQLString, gConn, transaction)
            If codtGRCV IsNot Nothing AndAlso codtGRCV.Rows.Count > 0 Then
                Dim POInsert As Int16 = 0

                For Each row As DataRow In codtGRCV.Rows
                    Dim txtCreated As Int16 = 0
                    'FOR BATCH_NO'
                    Dim dtBatchNo As New DataTable
                    SQLString = "Select REPLACE('WMSPO'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll(Max(Substring(BATCH_NO,6,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSPO1','WMSPO') as BATCH_NO From WMS_EBS_TRANS_ITX_PO_GRN"
                    dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "select RO_SEAL_NO from WMS_REPLENISH where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "'"
                    Dim codtTRPGR As New DataTable
                    codtTRPGR = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtTRPGR IsNot Nothing AndAlso codtTRPGR.Rows.Count > 0 Then
                        Dim dtACTION As New DataTable
                        SQLString = "Select TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,ACTION,STATUS,CREATION_DATE,LAST_UPDATE_DATE from EBS_WMS_TRANS_ITX_ACTION WHERE TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                        dtACTION = gDB.getDataTable(SQLString, gConn, transaction)

                        If dtACTION IsNot Nothing AndAlso dtACTION.Rows.Count > 0 Then

                            For Each rowACTION As DataRow In dtACTION.Rows
                                'for TRANSACTION_ID as use seq_no'
                                Dim dtSeq As New DataTable
                                SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                                dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

                                'FOR ACTION'
                                SQLString = "Select * from WMS_EBS_TRANS_ITX_ACTION where 1=2 " 'EBS_TRANSACTION_ID ='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                                Dim dtTRCode As New DataTable
                                dtTRCode = gDB.getDataTable(SQLString, gConn, transaction)

                                'If txtCreated = 0 Then
                                If dtTRCode Is Nothing Or dtTRCode.Rows.Count <= 0 Then
                                    'INSERT
                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,EBS_TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,STATUS,BATCH_NO,CREATION_DATE,LAST_UPDATE_DATE,CREATION_BY) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@EBS_TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@DOC_ID,@STATUS,@BATCH_NO,@CREATION_DATE,@LAST_UPDATE_DATE,@CREATION_BY)"
                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                    cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                    cmd.Parameters.AddWithValue("@ACTION", "NEW") 'rowACTION("ACTION").ToString.Trim
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@STATUS", "NEW")
                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                    'txtCreated = 1
                                Else
                                    'UPDATE
                                    sbCmdText = New StringBuilder()
                                    sbCmdText.Append("Update WMS_EBS_TRANS_ITX_ACTION Set ")
                                    sbCmdText.Append("ACTION = @ACTION,")
                                    sbCmdText.Append("IO_ID = @IO_ID,")
                                    sbCmdText.Append("COMPANY_ID = @COMPANY_ID,")
                                    sbCmdText.Append("DOC_TYPE = @DOC_TYPE,")
                                    sbCmdText.Append("DOC_ID = @DOC_ID,")
                                    sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                                    sbCmdText.Append(" Where EBS_TRANSACTION_ID = @EBS_TRANSACTION_ID")
                                    cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                    cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                End If

                            Next

                            SQLString = "Select top(1)ACTION,DOC_TYPE,EBS_TRANSACTION_ID,TRANSACTION_ID,BATCH_NO from WMS_EBS_TRANS_ITX_ACTION " &
                    "WHERE STATUS <> 'COMPLETED' and STATUS <> 'ERROR' and DOC_TYPE='PO' and EBS_TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "' order by TRANSACTION_ID desc"
                            Dim codtl As New DataTable
                            codtl = gDB.getDataTable(SQLString, gConn, transaction)

                            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then

                                If (codtl.Rows(0)("ACTION").ToString.Trim) = "NEW" And (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "PO" Then

                                    SQLString = "Select a.* FROM WMS_GOODSRCV_PA a Where a.GR_CODE ='" + row("GR_CODE") + "' "
                                    Dim codtGRCVD As New DataTable
                                    codtGRCVD = gDB.getDataTable(SQLString, gConn, transaction)
                                    If codtGRCVD IsNot Nothing AndAlso codtGRCVD.Rows.Count > 0 Then
                                        For Each rowGRCV As DataRow In codtGRCVD.Rows
                                            SQLString = "Select * FROM WMS_REPLENISH_D Where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "' and ROD_ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim & "' "
                                            Dim codtROD As New DataTable
                                            codtROD = gDB.getDataTable(SQLString, gConn, transaction)
                                            If codtROD IsNot Nothing AndAlso codtROD.Rows.Count > 0 Then
                                                'Location ID
                                                SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + rowGRCV("GRA_WH").ToString.Trim() + rowGRCV("GRA_LOC").ToString.Trim() + "' and IO_ID='" + rowGRCV("STORER_CODE").ToString.Trim() + "'"
                                                Dim dtLOC As New DataTable
                                                dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
                                                Dim LocationId As String = "0"
                                                If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                                                    LocationId = dtLOC.Rows(0)("LOCID").ToString()
                                                Else
                                                    SQLString = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists' WHERE EBS_TRANSACTION_ID='" & (codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                                    gDB.amendData(SQLString, gConn, transaction)
                                                End If

                                                'for seq_no'
                                                Dim dtTO As New DataTable
                                                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_PO_GRN"
                                                dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                                                SQLString = "Select * from EBS_WMS_TRANS_ITX_PO_DETAIL where TRANSACTION_ID='" & codtROD.Rows(0)("ROD_DOC_NO").ToString.Trim & "' and PO_HEADER_ID='" & codtROD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and PO_DISTRIBUTION_ID='" & codtROD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "' "
                                                Dim codt As New DataTable
                                                codt = gDB.getDataTable(SQLString, gConn, transaction)

                                                If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then

                                                    SQLString = "Select * from WMS_EBS_TRANS_ITX_PO_GRN where TRANSACTION_ID='" & codt.Rows(0)("TRANSACTION_ID").ToString.Trim & "' and PO_HEADER_ID='" & codt.Rows(0)("PO_HEADER_ID").ToString.Trim & "' and PO_DISTRIBUTION_ID='" & codt.Rows(0)("PO_DISTRIBUTION_ID").ToString.Trim & "' "
                                                    Dim codtPORGRN As New DataTable
                                                    codtPORGRN = gDB.getDataTable(SQLString, gConn, transaction)

                                                    If codtPORGRN Is Nothing Or codtPORGRN.Rows.Count <= 0 Then
                                                        'INSERT
                                                        SQLString = "Insert into WMS_EBS_TRANS_ITX_PO_GRN(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,PO_HEADER_ID,PO_DISTRIBUTION_ID,IO_ID,PO_LINE_ID,LINE_NUM,ITEM_ID,ITEM_NUMBER,ITEM_DESCRIPTION,NEED_BY_DATE,LINE_LOCATION_ID,UOM_CODE,UNIT_MEAS_LOOKUP_CODE,SHIP_TO_LOCATION_ID,QTY_RCV_TOLERANCE,SHIP_TO_ORGANIZATION_ID,SHIPMENT_NUM,RECEIVING_ROUTING_ID,QUANTITY_ORDERED,QUANTITY_DELIVERED,DISTRIBUTION_NUM,DELIVER_TO_LOCATION_ID,DEST_LOT_NUMBER,DEST_EXPIRY_DATE,DEST_SUBINVENTORY_CODE,DEST_QUANTITY,EBS_TRANSACTION_ID,DEST_LOCATOR_ID,DEST_UOM,CREATION_DATE,CREATION_BY,LAST_UPDATE_DATE,REMARK) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@PO_HEADER_ID,@PO_DISTRIBUTION_ID,@IO_ID,@PO_LINE_ID,@LINE_NUM,@ITEM_ID,@ITEN_NUMBER,@ITEM_DESCRIPTION,@NEED_BY_DATE,@LINE_LOCATION_ID,@UOM_CODE,@UNIT_MEAS_LOOKUP_CODE,@SHIP_TO_LOCATION_ID,@QTY_RCV_TOLERANCE,@SHIP_TO_ORGANIZATION_ID,@SHIPMENT_NUM,@RECEIVING_ROUTING_ID,@QUANTITY_ORDERED,@QUANTITY_DELIVERED,@DISTRIBUTION_NUM,@DELIVER_TO_LOCATION_ID,@DEST_LOT_NUMBER,@DEST_EXPIRY_DATE,@DEST_SUBINVENTORY_CODE,@DEST_QUANTITY,@EBS_TRANSACTION_ID,@DEST_LOCATOR_ID,@DEST_UOM,@CREATION_DATE,@CREATION_BY,@LAST_UPDATE_DATE,@REMARK)"
                                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                                        cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@BATCH_NO", codtl.Rows(0)("BATCH_NO").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtl.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@PO_HEADER_ID", codt.Rows(0)("PO_HEADER_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@PO_DISTRIBUTION_ID", codt.Rows(0)("PO_DISTRIBUTION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@IO_ID", codt.Rows(0)("IO_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@PO_LINE_ID", codt.Rows(0)("PO_LINE_ID").ToString.Trim())
                                                        cmd.Parameters.AddWithValue("@LINE_NUM", codt.Rows(0)("LINE_NUM").ToString.Trim())
                                                        cmd.Parameters.AddWithValue("@ITEM_ID", codt.Rows(0)("ITEM_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@ITEN_NUMBER", codt.Rows(0)("ITEM_NUMBER").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@ITEM_DESCRIPTION", codt.Rows(0)("ITEM_DESCRIPTION").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@NEED_BY_DATE", codt.Rows(0)("NEED_BY_DATE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@LINE_LOCATION_ID", codt.Rows(0)("LINE_LOCATION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@UOM_CODE", codt.Rows(0)("UOM_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@UNIT_MEAS_LOOKUP_CODE", codt.Rows(0)("UNIT_MEAS_LOOKUP_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@SHIP_TO_LOCATION_ID", codt.Rows(0)("SHIP_TO_LOCATION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@QTY_RCV_TOLERANCE", codt.Rows(0)("QTY_RCV_TOLERANCE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@SHIP_TO_ORGANIZATION_ID", codt.Rows(0)("SHIP_TO_ORGANIZATION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@SHIPMENT_NUM", codt.Rows(0)("SHIPMENT_NUM").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@RECEIVING_ROUTING_ID", codt.Rows(0)("RECEIVING_ROUTING_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@QUANTITY_ORDERED", codt.Rows(0)("QUANTITY_ORDERED").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@QUANTITY_DELIVERED", codt.Rows(0)("QUANTITY_DELIVERED").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DISTRIBUTION_NUM", codt.Rows(0)("DISTRIBUTION_NUM").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DELIVER_TO_LOCATION_ID", codt.Rows(0)("DELIVER_TO_LOCATION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_LOT_NUMBER", rowGRCV("GRA_BATCH_NO").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_EXPIRY_DATE", rowGRCV("GRA_EXPIRY_DATE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_SUBINVENTORY_CODE", rowGRCV("GRA_WH").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_LOCATOR_ID", LocationId)
                                                        cmd.Parameters.AddWithValue("@DEST_QUANTITY", rowGRCV("GRA_PA_QTY").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_UOM", codt.Rows(0)("UOM_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                        cmd.Parameters.AddWithValue("@CREATION_BY", row("SYS_CB").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                                        cmd.Parameters.AddWithValue("@REMARK", row("GR_TRACK_NO").ToString.Trim)
                                                        cmd.CommandType = System.Data.CommandType.Text
                                                        cmd.ExecuteScalar()
                                                        insertCount = insertCount + 1
                                                        POInsert = 1

                                                    End If

                                                    'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                                    SQLString = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (codt.Rows(0)("TRANSACTION_ID").ToString.Trim) & "'"
                                                    gDB.amendData(SQLString, gConn, transaction)
                                                End If

                                            End If

                                        Next

                                        SQLString = "UPDATE WMS_GOODSRCV SET EBS_UPDATE_STATUS ='Y' WHERE GR_CODE='" & row("GR_CODE").ToString.Trim & "' and GR_STATUS ='POSTED'"
                                        gDB.amendData(SQLString, gConn, transaction)
                                    End If


                                    'ElseIf (codtl.Rows(0)("ACTION").ToString.Trim) = "UPDATE" And (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "PO" Then
                                    '    'UPDATE
                                    '    'SQLString = "Select * from WMS_EBS_TRANS_ITX_PO_GRN " &
                                    '    '            "WHERE TRANSACTION_ID ='" & rowD("EBS_TRANSACTION_ID").ToString.Trim & "'"
                                    '    '        Dim dtROITXCode As New DataTable
                                    '    '        Dim adpROITXEBS As New SqlDataAdapter(SQLString, gConnEBS)
                                    '    '        adpROITXEBS.Fill(dtROITXCode)
                                    '    '        If dtROITXCode Is Nothing AndAlso dtROITXCode.Rows.Count <= 0 Then
                                    '    '            'FOR update status,error_msg and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                    '    '            UpdateSql = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE EBS_TRANSACTION_ID='" & (rowD("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                    '    '            Dim dtActions As New DataTable
                                    '    '            Dim adpEBSACTION As New SqlDataAdapter(UpdateSql, gConnEBS)
                                    '    '            adpEBSACTION.Fill(dtActions)
                                    '    '        End If

                                    '    SQLString = "Select * from WMS_REPLENISH_D WHERE ROD_PALLET_NO='" & rowH("PO_DISTRIBUTION_ID").ToString.Trim & "' and ROD_SERIES_NO='" & rowH("PO_HEADER_ID").ToString.Trim & "' and ROD_DOC_NO='" & rowD("EBS_TRANSACTION_ID").ToString.Trim & "'"
                                    '            Dim codtREPD As New DataTable
                                    '            codtREPD = gDB.getDataTable(SQLString)
                                    '            'For Each rowGRDD As DataRow In codtREPD.Rows
                                    '            SQLString = "Select * from WMS_EBS_TRANS_ITX_PO_GRN " &
                                    '    "WHERE TRANSACTION_ID ='" & codtREPD.Rows(0)("ROD_DOC_NO").ToString.Trim & "' and PO_HEADER_ID='" & codtREPD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and PO_DISTRIBUTION_ID='" & codtREPD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "'"
                                    '            Dim dtROCode As New DataTable
                                    '            Dim adpROEBS As New SqlDataAdapter(SQLString, gConnEBS)
                                    '            adpROEBS.Fill(dtROCode)
                                    '            If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                    '                'Update
                                    '                sbCmdText = New StringBuilder()
                                    '                sbCmdText.Append("Update WMS_EBS_TRANS_ITX_PO_GRN Set ")
                                    '                sbCmdText.Append("IO_ID = @IO_ID,")
                                    '                sbCmdText.Append("PO_LINE_ID = @PO_LINE_ID,")
                                    '                sbCmdText.Append("LINE_NUM = @LINE_NUM,")
                                    '                sbCmdText.Append("ITEM_ID = @ITEM_ID,")
                                    '                sbCmdText.Append("ITEN_NUMBER = @ITEN_NUMBER,")
                                    '                sbCmdText.Append("ITEM_DESCRIPTION = @ITEM_DESCRIPTION,")
                                    '                sbCmdText.Append("NEED_BY_DATE = @NEED_BY_DATE,")
                                    '                sbCmdText.Append("LINE_LOCATION_ID = @LINE_LOCATION_ID,")
                                    '                sbCmdText.Append("UOM_CODE = @UOM_CODE,")
                                    '                sbCmdText.Append("UNIT_MEAS_LOOKUP_CODE = @UNIT_MEAS_LOOKUP_CODE,")
                                    '                sbCmdText.Append("SHIP_TO_LOCATION_ID = @SHIP_TO_LOCATION_ID,")
                                    '                sbCmdText.Append("QTY_RCV_TOLERANCE = @QTY_RCV_TOLERANCE,")
                                    '                sbCmdText.Append("SHIP_TO_ORGANIZATION_ID = @SHIP_TO_ORGANIZATION_ID,")
                                    '                sbCmdText.Append("SHIPMENT_NUM = @SHIPMENT_NUM,")
                                    '                sbCmdText.Append("RECEIVING_ROUTING_ID = @RECEIVING_ROUTING_ID,")
                                    '                sbCmdText.Append("QUANTITY_ORDERED = @QUANTITY_ORDERED,")
                                    '                sbCmdText.Append("DISTRIBUTION_NUM = @DISTRIBUTION_NUM,")
                                    '                sbCmdText.Append("DELIVER_TO_LOCATION_ID = @DELIVER_TO_LOCATION_ID,")
                                    '                sbCmdText.Append("DEST_LOCATOR_ID = @DEST_LOCATOR_ID,")
                                    '                sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                                    '                sbCmdText.Append(" Where TRANSACTION_ID=@TRANSACTION_ID and PO_HEADER_ID=@PO_HEADER_ID and PO_DISTRIBUTION_ID=@PO_DISTRIBUTION_ID")
                                    '                cmd = New SqlCommand(sbCmdText.ToString(), gConnEBS)
                                    '                cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtREPD.Rows(0)("ROD_DOC_NO").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@PO_HEADER_ID", codtREPD.Rows(0)("ROD_SERIES_NO").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@PO_DISTRIBUTION_ID", codtREPD.Rows(0)("ROD_PALLET_NO").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@IO_ID", rowH("IO_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@PO_LINE_ID", rowH("PO_LINE_ID").ToString.Trim())
                                    '                cmd.Parameters.AddWithValue("@LINE_NUM", rowH("LINE_NUM").ToString.Trim())
                                    '                cmd.Parameters.AddWithValue("@ITEM_ID", rowH("ITEM_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@ITEN_NUMBER", rowH("ITEM_NUMBER").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@ITEM_DESCRIPTION", rowH("ITEM_DESCRIPTION").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@NEED_BY_DATE", rowH("NEED_BY_DATE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@LINE_LOCATION_ID", rowH("LINE_LOCATION_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@UOM_CODE", rowH("UOM_CODE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@UNIT_MEAS_LOOKUP_CODE", rowH("UNIT_MEAS_LOOKUP_CODE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@SHIP_TO_LOCATION_ID", rowH("SHIP_TO_LOCATION_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@QTY_RCV_TOLERANCE", rowH("QTY_RCV_TOLERANCE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@SHIP_TO_ORGANIZATION_ID", rowH("SHIP_TO_ORGANIZATION_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@SHIPMENT_NUM", rowH("SHIPMENT_NUM").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@RECEIVING_ROUTING_ID", rowH("RECEIVING_ROUTING_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@QUANTITY_ORDERED", rowH("QUANTITY_ORDERED").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@QUANTITY_DELIVERED", rowH("QUANTITY_DELIVERED").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@DISTRIBUTION_NUM", rowH("DISTRIBUTION_NUM").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@DELIVER_TO_LOCATION_ID", rowH("DELIVER_TO_LOCATION_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@DEST_LOCATOR_ID", "0")
                                    '                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                                    '                cmd.CommandType = System.Data.CommandType.Text
                                    '                cmd.ExecuteScalar()
                                    '                updateCount = updateCount + 1

                                    '                'FOR update status and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                    '                UpdateSql = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE EBS_TRANSACTION_ID='" & (rowGRDD("ROD_DOC_NO").ToString.Trim) & "'"
                                    '                Dim dtActions As New DataTable
                                    '                Dim adpEBSACTION As New SqlDataAdapter(UpdateSql, gConnEBS)
                                    '                adpEBSACTION.Fill(dtActions)

                                    '            End If


                                End If

                            Else
                                lblMSG.Text = "PO GRN already uploaded to EBS"
                            End If

                        End If

                    End If

                Next

                'For INTERFACE LOG
                If POInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "PO_GRN_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_PO_GRN_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_PO_GRN"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='PO'"
                    Dim POITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If POITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='PO' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "PO", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If

                End If

            End If

            transaction.Commit()

            'If lblMSG.Text = "" Then
            Dim strAlert As String = "<p>PO GRN Row Inserted : " & insertCount.ToString.Trim & "</p>"
            strAlert &= "<p>PO GRN Row Updated : " & updateCount.ToString.Trim & "</p>"
            'strAlert &= "<p>PO GRN data has been successfully imported!!</p>"
            lblMSG.Text = strAlert
            'Else
            '    lblMSG.Text = "PO_GRN data has been already updated"
            'End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import TO_GRN
    Public Sub ImportTO_GRNData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select a.* FROM WMS_GOODSRCV a Where a.GR_STATUS ='POSTED' and a.EBS_UPDATE_STATUS='N' and a.GR_DOC_TYPE='TO'"
            Dim codtGRCV As New DataTable
            codtGRCV = gDB.getDataTable(SQLString, gConn, transaction)
            If codtGRCV IsNot Nothing AndAlso codtGRCV.Rows.Count > 0 Then
                Dim TOInsert As Int16 = 0

                For Each row As DataRow In codtGRCV.Rows
                    Dim txtCreated As Int16 = 0

                    'FOR BATCH_NO'
                    Dim dtBatchNo As New DataTable
                    SQLString = "Select REPLACE('WMSTO'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,6,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSTO1','WMSTO') as BATCH_NO From WMS_EBS_TRANS_ITX_TO_GRN"
                    dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "select * from WMS_REPLENISH where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "'"
                    Dim codtTRPGR As New DataTable
                    codtTRPGR = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtTRPGR IsNot Nothing AndAlso codtTRPGR.Rows.Count > 0 Then
                        Dim dtACTION As New DataTable
                        SQLString = "Select TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,ACTION,STATUS,CREATION_DATE,LAST_UPDATE_DATE from EBS_WMS_TRANS_ITX_ACTION WHERE TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                        dtACTION = gDB.getDataTable(SQLString, gConn, transaction)
                        If dtACTION IsNot Nothing AndAlso dtACTION.Rows.Count > 0 Then
                            For Each rowACTION As DataRow In dtACTION.Rows
                                'for TRANSACTION_ID as use seq_no'
                                Dim dtSeq As New DataTable
                                SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                                dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

                                'FOR ACTION'
                                SQLString = "Select * from WMS_EBS_TRANS_ITX_ACTION where 1=2 " ' EBS_TRANSACTION_ID ='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                                Dim dtTRCode As New DataTable
                                dtTRCode = gDB.getDataTable(SQLString, gConn, transaction)

                                If dtTRCode Is Nothing Or dtTRCode.Rows.Count <= 0 Then
                                    'INSERT
                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,EBS_TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,STATUS,BATCH_NO,CREATION_DATE,LAST_UPDATE_DATE,CREATION_BY) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@EBS_TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@DOC_ID,@STATUS,@BATCH_NO,@CREATION_DATE,@LAST_UPDATE_DATE,@CREATION_BY)"
                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                    cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                    cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@STATUS", "NEW")
                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                    'txtCreated = 1
                                Else
                                    'UPDATE
                                    sbCmdText = New StringBuilder()
                                    sbCmdText.Append("Update WMS_EBS_TRANS_ITX_ACTION Set ")
                                    sbCmdText.Append("ACTION = @ACTION,")
                                    sbCmdText.Append("IO_ID = @IO_ID,")
                                    sbCmdText.Append("COMPANY_ID = @COMPANY_ID,")
                                    sbCmdText.Append("DOC_TYPE = @DOC_TYPE,")
                                    sbCmdText.Append("DOC_ID = @DOC_ID,")
                                    sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                                    sbCmdText.Append(" Where EBS_TRANSACTION_ID = @EBS_TRANSACTION_ID")
                                    cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                    cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                End If

                            Next

                            SQLString = "Select top(1) ACTION,DOC_TYPE,EBS_TRANSACTION_ID,TRANSACTION_ID,BATCH_NO from WMS_EBS_TRANS_ITX_ACTION " &
            "WHERE STATUS <> 'COMPLETED' and STATUS <> 'ERROR' and DOC_TYPE='TO' and EBS_TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "' order by TRANSACTION_ID desc"
                            Dim codtl As New DataTable
                            codtl = gDB.getDataTable(SQLString, gConn, transaction)

                            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                                If (codtl.Rows(0)("ACTION").ToString.Trim) = "NEW" And (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "TO" Then

                                    SQLString = "Select a.* FROM WMS_GOODSRCV_PA a  Where a.GR_CODE ='" + row("GR_CODE") + "' "
                                    Dim codtGRCVD As New DataTable
                                    codtGRCVD = gDB.getDataTable(SQLString, gConn, transaction)
                                    If codtGRCVD IsNot Nothing And codtGRCVD.Rows.Count > 0 Then
                                        For Each rowGRCV As DataRow In codtGRCVD.Rows
                                            SQLString = "Select * FROM WMS_REPLENISH_D Where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "' and ROD_ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim & "' and ROD_PALLET_NO='" & rowGRCV("GRA_PALLET_NO").ToString.Trim & "' "
                                            Dim codtROD As New DataTable
                                            codtROD = gDB.getDataTable(SQLString, gConn, transaction)
                                            If codtROD IsNot Nothing AndAlso codtROD.Rows.Count > 0 Then
                                                'for Location_Id'
                                                SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + rowGRCV("GRA_WH").ToString.Trim() + rowGRCV("GRA_LOC").ToString.Trim() + "' and IO_ID='" + rowGRCV("STORER_CODE").ToString.Trim() + "'"
                                                Dim dtLOC As New DataTable
                                                dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
                                                Dim LocationId As String = "0"
                                                If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                                                    LocationId = dtLOC.Rows(0)("LOCID").ToString()
                                                Else
                                                    SQLString = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists' WHERE EBS_TRANSACTION_ID='" & (codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                                    gDB.amendData(SQLString, gConn, transaction)
                                                End If

                                                'for seq_no'
                                                Dim dtTO As New DataTable
                                                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_TO_GRN"
                                                dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                                                SQLString = "Select * from EBS_WMS_TRANS_ITX_TO where TRANSACTION_ID='" & codtROD.Rows(0)("ROD_DOC_NO").ToString.Trim & "' and TO_HEADER_ID='" & codtROD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and LINE_ID='" & codtROD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "' and LINE_NUMBER='" & codtROD.Rows(0)("ROD_REF_NO").ToString.Trim & "'"
                                                Dim codt As New DataTable
                                                codt = gDB.getDataTable(SQLString, gConn, transaction)

                                                If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                                                    SQLString = "Select * from WMS_EBS_TRANS_ITX_TO_GRN where TRANSACTION_ID='" & codt.Rows(0)("TRANSACTION_ID").ToString.Trim & "' and TO_HEADER_ID='" & codt.Rows(0)("TO_HEADER_ID").ToString.Trim & "' and LINE_ID='" & codt.Rows(0)("LINE_ID").ToString.Trim & "' and LINE_NUMBER='" & codt.Rows(0)("LINE_NUMBER").ToString.Trim & "'"
                                                    Dim codtTOGRN As New DataTable
                                                    codtTOGRN = gDB.getDataTable(SQLString, gConn, transaction)
                                                    If codtTOGRN Is Nothing Or codtTOGRN.Rows.Count <= 0 Then
                                                        'INSERT
                                                        SQLString = "Insert into WMS_EBS_TRANS_ITX_TO_GRN(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,REQUEST_NUMBER,TO_HEADER_ID,LINE_ID,LINE_NUMBER,IO_ID,COMPANY_ID,ITEM_ID,ITEM_NUMBER,UOM_CODE,QUANTITY,LOT_NUMBER,FROM_SUBINVENTORY_CODE,FROM_LOCATOR_ID,TO_SUBINVENTORY_CODE,TO_LOCATOR_ID,MOVE_ORDER_TYPE,REVISION,DEST_SUBINVENTORY_CODE,DEST_LOCATION,DEST_QUANTITY,CREATION_BY,CREATION_DATE,LAST_UPDATE_DATE,EBS_TRANSACTION_ID,DEST_LOCATOR_ID) VALUES (@SEQ_NO, @BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@REQUEST_NUMBER,@TO_HEADER_ID,@LINE_ID,@LINE_NUMBER,@IO_ID,@COMPANY_ID,@ITEM_ID,@ITEM_NUMBER,@UOM_CODE,@QUANTITY,@LOT_NUMBER,@FROM_SUBINVENTORY_CODE,@FROM_LOCATOR_ID,@TO_SUBINVENTORY_CODE,@TO_LOCATOR_ID,@MOVE_ORDER_TYPE,@REVISION,@DEST_SUBINVENTORY_CODE,@DEST_LOCATION,@DEST_QUANTITY,@CREATION_BY,@CREATION_DATE,@LAST_UPDATE_DATE,@EBS_TRANSACTION_ID,@DEST_LOCATOR_ID)"
                                                        cmd = New SqlCommand(SQLString, gConn, transaction)
                                                        cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@BATCH_NO", codtl.Rows(0)("BATCH_NO").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtl.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@REQUEST_NUMBER", codt.Rows(0)("REQUEST_NUMBER").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TO_HEADER_ID", codt.Rows(0)("TO_HEADER_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@LINE_ID", codt.Rows(0)("LINE_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@LINE_NUMBER", codt.Rows(0)("LINE_NUMBER").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@IO_ID", codt.Rows(0)("IO_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@COMPANY_ID", codt.Rows(0)("COMPANY_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@ITEM_ID", codt.Rows(0)("ITEM_ID").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@ITEM_NUMBER", codt.Rows(0)("ITEM_NUMBER").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@UOM_CODE", codt.Rows(0)("UOM_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@QUANTITY", codt.Rows(0)("QUANTITY").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@LOT_NUMBER", codt.Rows(0)("LOT_NUMBER").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@FROM_SUBINVENTORY_CODE", codt.Rows(0)("FROM_SUBINVENTORY_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@FROM_LOCATOR_ID", IIf(String.IsNullOrEmpty(codt.Rows(0)("FROM_LOCATOR_ID").ToString.Trim), DBNull.Value, codt.Rows(0)("FROM_LOCATOR_ID").ToString.Trim))
                                                        cmd.Parameters.AddWithValue("@TO_SUBINVENTORY_CODE", codt.Rows(0)("TO_SUBINVENTORY_CODE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@TO_LOCATOR_ID", IIf(String.IsNullOrEmpty(codt.Rows(0)("TO_LOCATOR_ID").ToString.Trim), DBNull.Value, codt.Rows(0)("TO_LOCATOR_ID").ToString.Trim))
                                                        cmd.Parameters.AddWithValue("@MOVE_ORDER_TYPE", codt.Rows(0)("MOVE_ORDER_TYPE").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@REVISION", codt.Rows(0)("REVISION").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_SUBINVENTORY_CODE", rowGRCV("GRA_WH").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_QUANTITY", rowGRCV("GRA_PA_QTY").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@DEST_LOCATION", Left(Regex.Replace(rowGRCV("GRA_LOC").ToString.Trim, ".{2}", "$0."), 11))
                                                        cmd.Parameters.AddWithValue("@DEST_LOCATOR_ID", LocationId)
                                                        cmd.Parameters.AddWithValue("@CREATION_BY", row("SYS_CB").ToString.Trim)
                                                        cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                        cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                                        cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim)
                                                        cmd.CommandType = System.Data.CommandType.Text
                                                        cmd.ExecuteScalar()
                                                        insertCount = insertCount + 1
                                                        TOInsert = 1

                                                    End If

                                                End If

                                            End If

                                        Next

                                        SQLString = "UPDATE WMS_GOODSRCV SET EBS_UPDATE_STATUS ='Y' WHERE GR_CODE='" & row("GR_CODE").ToString.Trim & "' and GR_STATUS ='POSTED'"
                                        gDB.amendData(SQLString, gConn, transaction)

                                    End If

                                    '    ElseIf (rowD("ACTION").ToString.Trim) = "UPDATE" And (rowD("DOC_TYPE").ToString.Trim) = "TO" Then
                                    '        'UPDATE
                                    '        SQLString = "Select * from WMS_EBS_TRANS_ITX_TO_GRN " &
                                    '    "WHERE TRANSACTION_ID ='" & rowD("EBS_TRANSACTION_ID").ToString.Trim & "'"
                                    '        Dim dtROITXCode As New DataTable
                                    '        Dim adpROITXEBS As New SqlDataAdapter(SQLString, gConnEBS)
                                    '        adpROITXEBS.Fill(dtROITXCode)
                                    '        If dtROITXCode Is Nothing AndAlso dtROITXCode.Rows.Count <= 0 Then
                                    '            'FOR update status,error_msg and last_update_date in EBS_WMS_TRANS_ITX_ACTION'
                                    '            UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERR',ERROR_MESSAGE='Update must exist',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                    '            Dim dtActions As New DataTable
                                    '            Dim adpEBSACTION As New SqlDataAdapter(UpdateSql, gConnEBS)
                                    '            adpEBSACTION.Fill(dtActions)
                                    '        End If

                                    '        SQLString = "Select * from WMS_REPLENISH_D WHERE ROD_PALLET_NO='" & rowH("LINE_ID").ToString.Trim & "' and ROD_SERIES_NO='" & rowH("TO_HEADER_ID").ToString.Trim & "' and ROD_DOC_NO='" & rowD("EBS_TRANSACTION_ID").ToString.Trim & "' and ROD_REF_NO='" & rowH("LINE_NUMBER").ToString.Trim & "'"
                                    '        Dim codtREPD As New DataTable
                                    '        codtREPD = gDB.getDataTable(SQLString)
                                    '        For Each rowGRDD As DataRow In codtREPD.Rows
                                    '            SQLString = "Select * from WMS_EBS_TRANS_ITX_TO_GRN " &
                                    '"WHERE TRANSACTION_ID ='" & rowGRDD("ROD_DOC_NO").ToString.Trim & "' and TO_HEADER_ID='" & rowGRDD("ROD_SERIES_NO").ToString.Trim & "' and LINE_ID='" & rowGRDD("ROD_PALLET_NO").ToString.Trim & "' and LINE_NUMBER='" & rowGRDD("ROD_REF_NO").ToString.Trim & "'"
                                    '            Dim dtROCode As New DataTable
                                    '            Dim adpROEBS As New SqlDataAdapter(SQLString, gConnEBS)
                                    '            adpROEBS.Fill(dtROCode)
                                    '            If dtROCode IsNot Nothing AndAlso dtROCode.Rows.Count > 0 Then
                                    '                'Update
                                    '                sbCmdText = New StringBuilder()
                                    '                sbCmdText.Append("Update WMS_EBS_TRANS_ITX_TO_GRN Set ")
                                    '                sbCmdText.Append("REQUEST_NUMBER = @REQUEST_NUMBER,")
                                    '                sbCmdText.Append("IO_ID = @IO_ID,")
                                    '                sbCmdText.Append("COMPANY_ID = @COMPANY_ID,")
                                    '                sbCmdText.Append("ITEM_ID = @ITEM_ID,")
                                    '                sbCmdText.Append("ITEM_NUMBER = @ITEM_NUMBER,")
                                    '                sbCmdText.Append("UOM_CODE = @UOM_CODE,")
                                    '                sbCmdText.Append("QUANTITY = @QUANTITY,")
                                    '                sbCmdText.Append("LOT_NUMBER = @LOT_NUMBER,")
                                    '                sbCmdText.Append("FROM_SUBINVENTORY_CODE = @FROM_SUBINVENTORY_CODE,")
                                    '                sbCmdText.Append("FROM_LOCATOR_ID = @FROM_LOCATOR_ID,")
                                    '                sbCmdText.Append("TO_SUBINVENTORY_CODE = @TO_SUBINVENTORY_CODE,")
                                    '                sbCmdText.Append("TO_LOCATOR_ID = @TO_LOCATOR_ID,")
                                    '                sbCmdText.Append("MOVE_ORDER_TYPE = @MOVE_ORDER_TYPE,")
                                    '                sbCmdText.Append("REVISION = @REVISION,")
                                    '                sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                                    '                sbCmdText.Append(" Where TRANSACTION_ID=@TRANSACTION_ID and TO_HEADER_ID=@TO_HEADER_ID and LINE_ID=@LINE_ID and LINE_NUMBER=@LINE_NUMBER")
                                    '                cmd = New SqlCommand(sbCmdText.ToString(), gConnEBS)
                                    '                cmd.Parameters.AddWithValue("@TRANSACTION_ID", rowH("TRANSACTION_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@REQUEST_NUMBER", rowH("REQUEST_NUMBER").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@TO_HEADER_ID", rowH("TO_HEADER_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@LINE_ID", rowH("LINE_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@LINE_NUMBER", rowH("LINE_NUMBER").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@IO_ID", rowH("IO_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@COMPANY_ID", rowH("COMPANY_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@ITEM_ID", rowH("ITEM_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@ITEM_NUMBER", rowH("ITEM_NUMBER").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@UOM_CODE", rowH("UOM_CODE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@QUANTITY", rowH("QUANTITY").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@LOT_NUMBER", rowH("LOT_NUMBER").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@FROM_SUBINVENTORY_CODE", rowH("FROM_SUBINVENTORY_CODE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@FROM_LOCATOR_ID", rowH("FROM_LOCATOR_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@TO_SUBINVENTORY_CODE", rowH("TO_SUBINVENTORY_CODE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@TO_LOCATOR_ID", rowH("TO_LOCATOR_ID").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@MOVE_ORDER_TYPE", rowH("MOVE_ORDER_TYPE").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@REVISION", rowH("REVISION").ToString.Trim)
                                    '                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                                    '                cmd.CommandType = System.Data.CommandType.Text
                                    '                cmd.ExecuteScalar()
                                    '                updateCount = updateCount + 1

                                    '                'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                                    '                UpdateSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (rowD("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                    '                Dim dtActions As New DataTable
                                    '                Dim adpEBSACTION As New SqlDataAdapter(UpdateSql, gConnEBS)
                                    '                adpEBSACTION.Fill(dtActions)
                                    '            End If

                                    '        Next


                                End If

                            Else
                                lblMSG.Text = "TO GRN already uploaded to EBS"
                            End If

                            'FOR update status in EBS_WMS_TRANS_ITX_ACTION'
                            SQLString = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='COMPLETED',ERROR_MESSAGE='',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & (codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim) & "'"
                            gDB.amendData(SQLString, gConn, transaction)

                        End If

                    End If

                Next

                'For INTERFACE LOG
                If TOInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "TO_GRN_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_TO_GRN_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_TO_GRN"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='TO'"
                    Dim TOITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If TOITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='TO' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "TO", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            Dim strAlert As String = "<p>TO GRN Row Inserted : " & insertCount.ToString.Trim & "</p>"
            'strAlert &= "<p>TO GRN Row Updated : " & updateCount.ToString.Trim & "</p>"
            'strAlert &= "<p>TO GRN data has been successfully imported!!</p>"
            lblMSG.Text = strAlert

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import IPR_GRN
    Public Sub ImportIPR_GRNData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim UpdateSql As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountADJ As Int16 = 0
        Dim updateCountADJ As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select a.* FROM WMS_GOODSRCV a Where a.GR_STATUS ='POSTED' and a.EBS_UPDATE_STATUS='N' and a.GR_DOC_TYPE='IPR'"
            Dim codtGRCV As New DataTable
            codtGRCV = gDB.getDataTable(SQLString, gConn, transaction)
            If codtGRCV IsNot Nothing AndAlso codtGRCV.Rows.Count > 0 Then
                Dim IPRInsert As Int16 = 0

                For Each row As DataRow In codtGRCV.Rows
                    Dim txtCreated As Int16 = 0

                    SQLString = "select * from WMS_REPLENISH where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "'"
                    Dim codtTRPGR As New DataTable
                    codtTRPGR = gDB.getDataTable(SQLString, gConn, transaction)

                    If codtTRPGR IsNot Nothing AndAlso codtTRPGR.Rows.Count > 0 Then
                        'for Container_no'
                        Dim ROcontNo As String
                        If (codtTRPGR.Rows(0)("RO_CONTAINER_NO").ToString.Trim) <> "" Then
                            ROcontNo = codtTRPGR.Rows(0)("RO_CONTAINER_NO").ToString.Trim
                        Else
                            ROcontNo = ""
                        End If

                        Dim dtACTION As New DataTable
                        SQLString = "Select TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,ACTION,STATUS,CREATION_DATE,LAST_UPDATE_DATE from EBS_WMS_TRANS_ITX_ACTION WHERE TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                        dtACTION = gDB.getDataTable(SQLString, gConn, transaction)

                        If dtACTION IsNot Nothing AndAlso dtACTION.Rows.Count > 0 Then
                            'FOR BATCH_NO'
                            Dim dtBatchNo As New DataTable
                            SQLString = "Select REPLACE('WMSIPR'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSIPR1','WMSIPR') as BATCH_NO From WMS_EBS_TRANS_ITX_IPR_GRN"
                            dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                            'FOR stock Adjustment BATCH_NO'
                            Dim dtADJBatchNo As New DataTable
                            SQLString = "Select REPLACE('WMSLOT'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSLOT1','WMSLOT') as BATCH_NO From WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
                            dtADJBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                            Dim dtSeq As New DataTable
                            Dim dtstkAdjSeq As New DataTable
                            Dim dtTRCode As New DataTable
                            For Each rowACTION As DataRow In dtACTION.Rows
                                'FOR ACTION -- Changed this query to always return 0 records in order to generate new action every time'
                                SQLString = "Select * from WMS_EBS_TRANS_ITX_ACTION where 1=2 and EBS_TRANSACTION_ID ='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "'"
                                dtTRCode = New DataTable
                                dtTRCode = gDB.getDataTable(SQLString, gConn, transaction)

                                'for TRANSACTION_ID as use seq_no'
                                dtSeq = New DataTable
                                SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                                dtSeq = gDB.getDataTable(SQLString, gConn, transaction)
                                If dtTRCode Is Nothing Or dtTRCode.Rows.Count <= 0 Then
                                    'INSERT
                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,EBS_TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,STATUS,BATCH_NO,RELATED_NO,CREATION_DATE,LAST_UPDATE_DATE,CREATION_BY) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@EBS_TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@DOC_ID,@STATUS,@BATCH_NO,@RELATED_NO,@CREATION_DATE,@LAST_UPDATE_DATE,@CREATION_BY)"
                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                    cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                    cmd.Parameters.AddWithValue("@ACTION", "NEW")
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", dtACTION.Rows(0)("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", dtACTION.Rows(0)("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", dtACTION.Rows(0)("DOC_TYPE").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_ID", dtACTION.Rows(0)("DOC_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@STATUS", "TEMP")
                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@RELATED_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()
                                    'txtCreated = 1

                                    'FOR stock Adjustment
                                    'for TRANSACTION_ID as use seq_no'
                                    dtstkAdjSeq = New DataTable
                                    SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                                    dtstkAdjSeq = gDB.getDataTable(SQLString, gConn, transaction)

                                    'INSERT
                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,IO_ID,DOC_TYPE,STATUS,BATCH_NO,COMPANY_ID,EBS_TRANSACTION_ID,CREATION_DATE,LAST_UPDATE_DATE,CREATION_BY) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@IO_ID,@DOC_TYPE,@STATUS,@BATCH_NO,@COMPANY_ID,@EBS_TRANSACTION_ID,@CREATION_DATE,@LAST_UPDATE_DATE,@CREATION_BY)"
                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                    cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                    cmd.Parameters.AddWithValue("@ACTION", "NEW")
                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtstkAdjSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@IO_ID", dtACTION.Rows(0)("IO_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@DOC_TYPE", "IPR_ADJ")
                                    cmd.Parameters.AddWithValue("@COMPANY_ID", dtACTION.Rows(0)("COMPANY_ID").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@STATUS", "TEMP")
                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                    cmd.CommandType = System.Data.CommandType.Text
                                    cmd.ExecuteScalar()

                                End If

                            Next

                            SQLString = "Select ACTION,DOC_TYPE,EBS_TRANSACTION_ID,TRANSACTION_ID,BATCH_NO from WMS_EBS_TRANS_ITX_ACTION " &
                               "WHERE  DOC_TYPE='IPR' and EBS_TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "' order by TRANSACTION_ID desc"
                            Dim codtl As New DataTable
                            codtl = gDB.getDataTable(SQLString, gConn, transaction)

                            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then

                                If (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "IPR" Then '(codtl.Rows(0)("ACTION").ToString.Trim) = "NEW" And 
                                    SQLString = "Select a.* FROM WMS_GOODSRCV_PA a Where a.GR_CODE ='" + row("GR_CODE") + "' and a.GRA_IPR_LOT_OUTSTANDING > 0  "
                                    Dim codtGRCVD As New DataTable
                                    codtGRCVD = gDB.getDataTable(SQLString, gConn, transaction)
                                    If codtGRCVD IsNot Nothing AndAlso codtGRCVD.Rows.Count > 0 Then
                                        For Each rowGRCV As DataRow In codtGRCVD.Rows
                                            'For STORER 
                                            Dim IO_CODE As String
                                            SQLString = "select Top(1) IO_CODE from EBS_WMS_COMPANY_MASTER where IO_ID = '" & rowGRCV("STORER_CODE").ToString.Trim() & "' "
                                            Dim dtIOCode As DataTable
                                            dtIOCode = gDB.getDataTable(SQLString, gConn, transaction)
                                            If dtIOCode IsNot Nothing AndAlso dtIOCode.Rows.Count > 0 Then
                                                IO_CODE = dtIOCode.Rows(0)("IO_CODE").ToString.Trim
                                            Else
                                                IO_CODE = ""
                                            End If

                                            SQLString = "Select * FROM WMS_REPLENISH_D Where RO_CODE='" & row("GR_DOC_NO").ToString.Trim & "' and ROD_ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim & "'"
                                            Dim codtROD As New DataTable
                                            codtROD = gDB.getDataTable(SQLString, gConn, transaction)
                                            If codtROD IsNot Nothing AndAlso codtROD.Rows.Count > 0 Then

                                                'FOR EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT'
                                                SQLString = "Select * from EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT where IPR_HEADER_ID='" & codtROD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and IPR_LINE_ID='" & codtROD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "' and REMAINING_QUANTITY>0 "
                                                Dim codtLot As New DataTable
                                                codtLot = gDB.getDataTable(SQLString, gConn, transaction)

                                                If codtLot IsNot Nothing AndAlso codtLot.Rows.Count > 0 Then
                                                    SQLString = "Select SUM(REMAINING_QUANTITY) AS Total_Quantity from EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT where IPR_HEADER_ID='" & codtROD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and IPR_LINE_ID='" & codtROD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "'"
                                                    Dim dtTtlQTY As New DataTable
                                                    dtTtlQTY = gDB.getDataTable(SQLString, gConn, transaction)

                                                    If dtTtlQTY.Rows(0)("Total_Quantity").ToString.Trim > 0 Then
                                                        'for seq_no'


                                                        SQLString = "Select * from EBS_WMS_TRANS_ITX_IPR_DETAIL where TRANSACTION_ID='" & codtROD.Rows(0)("ROD_DOC_NO").ToString.Trim & "' and IPR_HEADER_ID='" & codtROD.Rows(0)("ROD_SERIES_NO").ToString.Trim & "' and IPR_LINE_ID='" & codtROD.Rows(0)("ROD_PALLET_NO").ToString.Trim & "' "
                                                        Dim codt As New DataTable
                                                        codt = gDB.getDataTable(SQLString, gConn, transaction)

                                                        Dim DESTQty As Double
                                                        DESTQty = codtLot.Rows(0)("TRANSACTION_QUANTITY").ToString.Trim

                                                        If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                                                            'SQLString = "Select * from WMS_EBS_TRANS_ITX_IPR_GRN where TRANSACTION_ID='" & codt.Rows(0)("TRANSACTION_ID").ToString.Trim & "' and IPR_HEADER_ID='" & codt.Rows(0)("IPR_HEADER_ID").ToString.Trim & "' and IPR_LINE_ID='" & codt.Rows(0)("IPR_LINE_ID").ToString.Trim & "' "
                                                            'Dim codtIPRGRN As New DataTable
                                                            'codtIPRGRN = gDB.getDataTable(SQLString, gConn, transaction)

                                                            'If codtIPRGRN Is Nothing Or codtIPRGRN.Rows.Count <= 0 Then
                                                            'INSERT

                                                            'End If


                                                            'for LOCATOR'
                                                            Dim LOC As String
                                                            Dim location As String
                                                            LOC = Right((rowGRCV("GRA_REMARK").ToString.Trim()), 8)
                                                            If LOC <> "" Then
                                                                location = LOC.Substring(0, 2) + "." + LOC.Substring(2, 2) + "." + LOC.Substring(4, 2) + "." + LOC.Substring(6, 2)
                                                            Else
                                                                location = ""
                                                            End If

                                                            'For UOM_CODE
                                                            SQLString = "select Top(1) ITM_UOM as PRIMARY_UOM_CODE from WMS_ITEM where ITM_CODE = '" & codt.Rows(0)("ITEM_ID").ToString.Trim() & "' and STORER_CODE='" & rowGRCV("STORER_CODE").ToString.Trim() & "' and ITM_SKU_NO='" & codt.Rows(0)("ITEM_NUMBER").ToString.Trim & "'"
                                                            Dim dtUOM_CODE As DataTable
                                                            dtUOM_CODE = gDB.getDataTable(SQLString, gConn, transaction)
                                                            Dim UOM_CODE As String
                                                            If dtUOM_CODE IsNot Nothing AndAlso dtUOM_CODE.Rows.Count > 0 Then
                                                                UOM_CODE = dtUOM_CODE.Rows(0)("PRIMARY_UOM_CODE").ToString.Trim
                                                            Else
                                                                UOM_CODE = ""
                                                            End If

                                                            Dim RCVQTY_BALANCE As Double
                                                            RCVQTY_BALANCE = (rowGRCV("GRA_IPR_LOT_OUTSTANDING").ToString.Trim())
                                                            Dim codtStkAdj As New DataTable
                                                            Dim dtTO As New DataTable
                                                            SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + rowGRCV("GRA_REMARK").ToString.Trim() + "' and IO_ID='" + rowGRCV("STORER_CODE").ToString.Trim() + "'"
                                                            Dim dtLOC As New DataTable
                                                            dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
                                                            Dim LocationId As String = "0"
                                                            If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                                                                LocationId = dtLOC.Rows(0)("LOCID").ToString()
                                                            Else
                                                                If (dtTRCode IsNot Nothing And dtTRCode.Rows.Count > 0) Then
                                                                    UpdateSql = "Update WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists'  WHERE TRANSACTION_ID in (Select TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION where EBS_TRANSACTION_ID ='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "') "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)
                                                                End If

                                                                If (dtstkAdjSeq IsNot Nothing And dtstkAdjSeq.Rows.Count > 0) Then
                                                                    SQLString = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists' WHERE TRANSACTION_ID in ('" & (dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim) & "','" & (dtstkAdjSeq.Rows(0)("TRANSACTION_ID").ToString.Trim) & "')"
                                                                    gDB.amendData(SQLString, gConn, transaction)
                                                                End If

                                                            End If
                                                            For Each rowDtlLot As DataRow In codtLot.Rows
                                                                'FOR SEQ NO.'

                                                                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_IPR_GRN"
                                                                dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                                                                Dim dtADJSEQ As New DataTable
                                                                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
                                                                dtADJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

                                                                ' creating record for WMS_EBS_TRANS_ITX_IPR_GRN
                                                                SQLString = "Insert into WMS_EBS_TRANS_ITX_IPR_GRN(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,IPR_HEADER_ID,IPR_NO,IPR_LINE_ID,IPR_LINE_NUMBER,TO_PERSON_ID,ITEM_ID,ITEM_NUMBER,IPR_UOM_CODE,QUANTITY,SOURCE_IO_ID,DEST_IO_ID,CATEGORY_ID,DEST_TYPE_CODE,DEST_QUANTITY,REMARK,CREATION_BY,CREATION_DATE,LAST_UPDATE_DATE,EBS_TRANSACTION_ID,REQ_DISTRIBUTION_ID,LOT_NUMBER,SHIPMENT_LINE_ID,FROM_IO_ID,TO_IO_ID,SHIPMENT_IO_ID,SHIPMENT_NUM,QUANTITY_SHIPPED,QUANTITY_RECEIVED,SOURCE_DOCUMENT_CODE,ROUTING_HEADER_ID,DELIVER_TO_PERSON_ID,DELIVER_TO_LOCATION_ID,SHIP_TO_LOCATION_ID,TRANSACTION_QUANTITY,ORIGINATION_DATE) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@IPR_HEADER_ID,@IPR_NO,@IPR_LINE_ID,@IPR_LINE_NUMBER,@TO_PERSON_ID,@ITEM_ID,@ITEM_NUMBER,@IPR_UOM_CODE,@QUANTITY,@SOURCE_IO_ID,@DEST_IO_ID,@CATEGORY_ID,@DEST_TYPE_CODE,@DEST_QUANTITY,@REMARK,@CREATION_BY,@CREATION_DATE,@LAST_UPDATE_DATE,@EBS_TRANSACTION_ID,@REQ_DISTRIBUTION_ID,@LOT_NUMBER,@SHIPMENT_LINE_ID,@FROM_IO_ID,@TO_IO_ID,@SHIPMENT_IO_ID,@SHIPMENT_NUM,@QUANTITY_SHIPPED,@QUANTITY_RECEIVED,@SOURCE_DOCUMENT_CODE,@ROUTING_HEADER_ID,@DELIVER_TO_PERSON_ID,@DELIVER_TO_LOCATION_ID,@SHIP_TO_LOCATION_ID,@TRANSACTION_QUANTITY,@ORIGINATION_DATE)"
                                                                cmd = New SqlCommand(SQLString, gConn, transaction)
                                                                cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@BATCH_NO", codtl.Rows(0)("BATCH_NO").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtl.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@IPR_HEADER_ID", codt.Rows(0)("IPR_HEADER_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@IPR_NO", codt.Rows(0)("IPR_NO").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@IPR_LINE_ID", codt.Rows(0)("IPR_LINE_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@IPR_LINE_NUMBER", codt.Rows(0)("IPR_LINE_NUMBER").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@TO_PERSON_ID", codt.Rows(0)("TO_PERSON_ID").ToString.Trim())
                                                                cmd.Parameters.AddWithValue("@ITEM_ID", codt.Rows(0)("ITEM_ID").ToString.Trim())
                                                                cmd.Parameters.AddWithValue("@ITEM_NUMBER", codt.Rows(0)("ITEM_NUMBER").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@IPR_UOM_CODE", codt.Rows(0)("IPR_UOM_CODE").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@QUANTITY", codt.Rows(0)("QUANTITY").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@SOURCE_IO_ID", codt.Rows(0)("SOURCE_IO_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@DEST_IO_ID", codt.Rows(0)("DEST_IO_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@CATEGORY_ID", codt.Rows(0)("CATEGORY_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@DEST_TYPE_CODE", codt.Rows(0)("DEST_TYPE_CODE").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@REQ_DISTRIBUTION_ID", rowDtlLot("REQ_DISTRIBUTION_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@DEST_QUANTITY", rowDtlLot("TRANSACTION_QUANTITY").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@REMARK", row("GR_TRACK_NO").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@CREATION_BY", row("SYS_CB").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                                                cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@LOT_NUMBER", rowDtlLot("LOT_NUMBER").ToString.Trim)


                                                                cmd.Parameters.AddWithValue("@SHIPMENT_LINE_ID", rowDtlLot("SHIPMENT_LINE_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@FROM_IO_ID", rowDtlLot("FROM_IO_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@TO_IO_ID", rowDtlLot("TO_IO_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@SHIPMENT_IO_ID", rowDtlLot("SHIPMENT_IO_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@SHIPMENT_NUM", rowDtlLot("SHIPMENT_NUM").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@QUANTITY_SHIPPED", rowDtlLot("QUANTITY_SHIPPED").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@QUANTITY_RECEIVED", rowDtlLot("QUANTITY_RECEIVED").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@SOURCE_DOCUMENT_CODE", rowDtlLot("SOURCE_DOCUMENT_CODE").ToString.Trim)

                                                                cmd.Parameters.AddWithValue("@ROUTING_HEADER_ID", rowDtlLot("ROUTING_HEADER_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@DELIVER_TO_PERSON_ID", rowDtlLot("DELIVER_TO_PERSON_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@DELIVER_TO_LOCATION_ID", rowDtlLot("DELIVER_TO_LOCATION_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@SHIP_TO_LOCATION_ID", rowDtlLot("SHIP_TO_LOCATION_ID").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@TRANSACTION_QUANTITY", rowDtlLot("TRANSACTION_QUANTITY").ToString.Trim)
                                                                cmd.Parameters.AddWithValue("@ORIGINATION_DATE", rowDtlLot("ORIGINATION_DATE").ToString.Trim)

                                                                cmd.CommandType = System.Data.CommandType.Text
                                                                cmd.ExecuteScalar()
                                                                insertCount = insertCount + 1
                                                                IPRInsert = 1

                                                                'update WMS_EBS_TRANS_ITX_ACTION Status'
                                                                If (dtTRCode IsNot Nothing And dtTRCode.Rows.Count > 0) Then
                                                                    UpdateSql = "Update WMS_EBS_TRANS_ITX_ACTION SET ACTION='NEW', STATUS='NEW' WHERE TRANSACTION_ID in (Select TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION where EBS_TRANSACTION_ID ='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "' and STATUS='TEMP') "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)
                                                                End If

                                                                If (dtstkAdjSeq IsNot Nothing And dtstkAdjSeq.Rows.Count > 0) Then
                                                                    UpdateSql = "Update WMS_EBS_TRANS_ITX_ACTION SET ACTION='NEW', STATUS='NEW' WHERE TRANSACTION_ID in ('" & dtstkAdjSeq.Rows(0)("TRANSACTION_ID").ToString.Trim & "','" & dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim & "') and  STATUS='TEMP' "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)
                                                                End If



                                                                SQLString = "Select * from WMS_EBS_TRANS_ITX_ACTION " &
                                       "WHERE  DOC_TYPE='IPR_ADJ' and EBS_TRANSACTION_ID='" & codtTRPGR.Rows(0)("RO_SEAL_NO").ToString.Trim & "' order by TRANSACTION_ID desc"

                                                                codtStkAdj = gDB.getDataTable(SQLString, gConn, transaction)


                                                                If (RCVQTY_BALANCE <= (rowDtlLot("REMAINING_QUANTITY").ToString.Trim)) Then
                                                                    Dim shippedQty, diffQty As Double
                                                                    shippedQty = (rowDtlLot("QUANTITY_SHIPPED").ToString.Trim())
                                                                    diffQty = (rowDtlLot("REMAINING_QUANTITY").ToString.Trim() - RCVQTY_BALANCE)

                                                                    'update EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT'
                                                                    UpdateSql = "Update EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT SET " &
                                                            "REMAINING_QUANTITY=" & diffQty & " " &
                                                            "where IPR_HEADER_ID = '" & rowDtlLot("IPR_HEADER_ID").ToString.Trim & "' " &
                                                            "and IPR_LINE_ID = '" & rowDtlLot("IPR_LINE_ID").ToString.Trim & "' " &
                                                            "and SHIPMENT_HEADER_ID = '" & rowDtlLot("SHIPMENT_HEADER_ID").ToString.Trim & "' " &
                                                            "and SHIPMENT_LINE_ID = '" & rowDtlLot("SHIPMENT_LINE_ID").ToString.Trim & "' " &
                                                            "and TRANSACTION_ID = '" & rowDtlLot("TRANSACTION_ID").ToString.Trim & "' "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)


                                                                    'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                                                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,SUBINVENTORY_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@SUBINVENTORY_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY)"
                                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                                    cmd.Parameters.AddWithValue("@SEQ_NO", dtADJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtStkAdj.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "I")
                                                                    cmd.Parameters.AddWithValue("@DOC_TYPE", "IPR_ADJ")
                                                                    cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                                                                    cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", row("GR_WH_CODE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@LOCATOR", DBNull.Value)
                                                                    cmd.Parameters.AddWithValue("@LOCATION_ID", DBNull.Value)
                                                                    cmd.Parameters.AddWithValue("@LOT_NUMBER", rowDtlLot("LOT_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ITEM_NUMBER", rowDtlLot("ITEM_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@UOM_CODE", UOM_CODE)
                                                                    cmd.Parameters.AddWithValue("@QUANTITY", RCVQTY_BALANCE)
                                                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", Session("usr_nickname"))
                                                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                                                    cmd.CommandType = System.Data.CommandType.Text
                                                                    cmd.ExecuteScalar()
                                                                    insertCountADJ = insertCountADJ + 1


                                                                    UpdateSql = "update WMS_EBS_TRANS_ITX_IPR_GRN set DEST_QUANTITY=" & RCVQTY_BALANCE & " where SEQ_NO='" & dtTO.Rows(0)("SEQ_NO").ToString.Trim & "' "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)

                                                                    RCVQTY_BALANCE = 0
                                                                    Exit For

                                                                Else

                                                                    'update EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT'
                                                                    UpdateSql = "Update EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT SET " &
                                                            "REMAINING_QUANTITY='0' " &
                                                            "where IPR_HEADER_ID = '" & rowDtlLot("IPR_HEADER_ID").ToString.Trim & "' " &
                                                            "and IPR_LINE_ID = '" & rowDtlLot("IPR_LINE_ID").ToString.Trim & "' " &
                                                            "and SHIPMENT_HEADER_ID = '" & rowDtlLot("SHIPMENT_HEADER_ID").ToString.Trim & "' " &
                                                            "and SHIPMENT_LINE_ID = '" & rowDtlLot("SHIPMENT_LINE_ID").ToString.Trim & "' " &
                                                            "and TRANSACTION_ID = '" & rowDtlLot("TRANSACTION_ID").ToString.Trim & "' "
                                                                    gDB.amendData(UpdateSql, gConn, transaction)

                                                                    DESTQty = rowDtlLot("TRANSACTION_QUANTITY").ToString.Trim
                                                                    'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                                                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,SUBINVENTORY_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@SUBINVENTORY_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY)"
                                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                                    cmd.Parameters.AddWithValue("@SEQ_NO", dtADJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtStkAdj.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "I")
                                                                    cmd.Parameters.AddWithValue("@DOC_TYPE", "IPR_ADJ")
                                                                    cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                                                                    cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", row("GR_WH_CODE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@LOCATOR", DBNull.Value)
                                                                    cmd.Parameters.AddWithValue("@LOCATION_ID", DBNull.Value)
                                                                    cmd.Parameters.AddWithValue("@LOT_NUMBER", rowDtlLot("LOT_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ITEM_NUMBER", rowDtlLot("ITEM_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@UOM_CODE", UOM_CODE)
                                                                    cmd.Parameters.AddWithValue("@QUANTITY", rowDtlLot("REMAINING_QUANTITY").ToString.Trim())
                                                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                                                    cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", Session("usr_nickname"))
                                                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", rowGRCV("SYS_LUD").ToString.Trim)
                                                                    cmd.CommandType = System.Data.CommandType.Text
                                                                    cmd.ExecuteScalar()
                                                                    insertCountADJ = insertCountADJ + 1

                                                                    RCVQTY_BALANCE = (RCVQTY_BALANCE - (rowDtlLot("REMAINING_QUANTITY").ToString.Trim))

                                                                End If
                                                                'DESTQty = (rowGRCV("GRA_IPR_LOT_OUTSTANDING").ToString.Trim()) - RCVQTY_BALANCE
                                                                DESTQty = rowDtlLot("REMAINING_QUANTITY").ToString.Trim
                                                                UpdateSql = "update WMS_EBS_TRANS_ITX_IPR_GRN set DEST_QUANTITY=" & DESTQty & " where SEQ_NO='" & dtTO.Rows(0)("SEQ_NO").ToString.Trim & "' "
                                                                gDB.amendData(UpdateSql, gConn, transaction)
                                                            Next

                                                            If (RCVQTY_BALANCE <= 0) Then
                                                                'for GRA_IPR_LOT_OUTSTANDING'
                                                                UpdateSql = "Update WMS_GOODSRCV_PA SET " &
                                                            "GRA_IPR_LOT_OUTSTANDING = 0 " &
                                                            "where GR_CODE = '" & rowGRCV("GR_CODE").ToString.Trim & "' and GRA_SEQ= '" & rowGRCV("GRA_SEQ").ToString.Trim & "' and  GRA_ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim & "'"
                                                                gDB.amendData(UpdateSql, gConn, transaction)
                                                            Else
                                                                UpdateSql = "Update WMS_GOODSRCV_PA SET " &
                                                          "GRA_IPR_LOT_OUTSTANDING = " & RCVQTY_BALANCE & " " &
                                                          "where GR_CODE = '" & rowGRCV("GR_CODE").ToString.Trim & "' and GRA_SEQ= '" & rowGRCV("GRA_SEQ").ToString.Trim & "' and  GRA_ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim & "' "
                                                                gDB.amendData(UpdateSql, gConn, transaction)
                                                            End If



                                                            If RCVQTY_BALANCE > 0 Then
                                                                Dim GRALOTOUTTADINGQty As Double
                                                                GRALOTOUTTADINGQty = rowGRCV("GRA_IPR_LOT_OUTSTANDING").ToString.Trim()
                                                                GRALOTOUTTADINGQty = RCVQTY_BALANCE
                                                            End If

                                                            'FOR SEQ NO.'
                                                            Dim dtADJJSEQ As New DataTable
                                                            SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
                                                            dtADJJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

                                                            'Get SKU Number for Item code
                                                            SQLString = "select TOP(1) ITM_SKU_NO from WMS_ITEM where ITM_CODE='" & rowGRCV("GRA_ITM_CODE").ToString.Trim() & "' and STORER_CODE='" & rowGRCV("STORER_CODE").ToString.Trim() & "'"
                                                            Dim dtItemSKCODE As DataTable
                                                            dtItemSKCODE = gDB.getDataTable(SQLString, gConn, transaction)


                                                            'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                                                            SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,SUBINVENTORY_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY,EXPIRY_DATE) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@SUBINVENTORY_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY,@EXPIRY_DATE)"
                                                            cmd = New SqlCommand(SQLString, gConn, transaction)
                                                            cmd.Parameters.AddWithValue("@SEQ_NO", dtADJJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtStkAdj.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@TRANSACTION_DATE", row("GR_DATE").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "R")
                                                            cmd.Parameters.AddWithValue("@DOC_TYPE", "IPR_ADJ")
                                                            cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                                                            cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", rowGRCV("GRA_REMARK").ToString.Trim.Remove(rowGRCV("GRA_REMARK").ToString.Length - 8, 8))
                                                            cmd.Parameters.AddWithValue("@LOCATOR", location)
                                                            cmd.Parameters.AddWithValue("@LOCATION_ID", LocationId)
                                                            cmd.Parameters.AddWithValue("@LOT_NUMBER", rowGRCV("GRA_BATCH_NO").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@ITEM_NUMBER", dtItemSKCODE.Rows(0)("ITM_SKU_NO").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@UOM_CODE", UOM_CODE)
                                                            cmd.Parameters.AddWithValue("@QUANTITY", rowGRCV("GRA_IPR_LOT_OUTSTANDING").ToString.Trim - RCVQTY_BALANCE)
                                                            cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                                            cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                                            cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", Session("usr_nickname"))
                                                            cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", rowGRCV("SYS_LUD").ToString.Trim)
                                                            cmd.Parameters.AddWithValue("@EXPIRY_DATE", rowGRCV("GRA_EXPIRY_DATE").ToString.Trim)
                                                            cmd.CommandType = System.Data.CommandType.Text
                                                            cmd.ExecuteScalar()
                                                            insertCountADJ = insertCountADJ + 1

                                                            'For ITEM_CODE'
                                                            SQLString = "select TOP(1) ITM_CODE from WMS_ITEM where ITM_SKU_NO = '" & dtItemSKCODE.Rows(0)("ITM_SKU_NO").ToString.Trim & "' and STORER_CODE='" & rowGRCV("STORER_CODE").ToString.Trim() & "' "
                                                            Dim dtITMCODE As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                            Dim ITM_CODE As String
                                                            If dtITMCODE IsNot Nothing AndAlso dtITMCODE.Rows.Count > 0 Then
                                                                ITM_CODE = dtITMCODE.Rows(0)("ITM_CODE").ToString.Trim
                                                            Else
                                                                ITM_CODE = ""
                                                            End If

                                                            Dim stklocation As String
                                                            stklocation = location.Replace(".", "").ToString()

                                                            'For Expiry_date & Org_Qty
                                                            SQLString = "select ILOC_EXPIRY_DATE,ILOC_BAL_QTY from WMS_ITEM_LOC_BAL where ITM_CODE = '" & ITM_CODE & "' and STORER_CODE='" & rowGRCV("STORER_CODE").ToString.Trim() & "' and ILOC_BATCH_NO='" & rowGRCV("GRA_BATCH_NO").ToString.Trim & "' and ILOC_LOC='" & rowGRCV("GRA_LOC").ToString.Trim & "' and ILOC_WH='" & rowGRCV("GRA_WH").ToString.Trim & "'"
                                                            Dim dtITMBALSTK As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                                                            Dim ExpiryDate, BAL_QTY As String
                                                            If dtITMBALSTK IsNot Nothing AndAlso dtITMBALSTK.Rows.Count > 0 Then
                                                                ExpiryDate = dtITMBALSTK.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim
                                                                BAL_QTY = rowGRCV("GRA_IPR_LOT_OUTSTANDING").ToString.Trim - RCVQTY_BALANCE 'rowGRCV("GRA_PA_QTY").ToString.Trim
                                                            Else
                                                                ExpiryDate = ""
                                                                BAL_QTY = ""
                                                            End If


                                                            INSPOSTDATASR(rowGRCV("STORER_CODE").ToString.Trim(), rowGRCV("GRA_BATCH_NO").ToString.Trim, rowGRCV("GRA_BATCH_NO").ToString.Trim, rowGRCV("GRA_LOC").ToString.Trim, stklocation, "FGTM", rowGRCV("GRA_REMARK").ToString.Trim.Remove(rowGRCV("GRA_REMARK").ToString.Length - 8, 8), ITM_CODE, BAL_QTY, BAL_QTY, "000", "000", ExpiryDate, transaction, gConn)


                                                        End If

                                                    End If

                                                End If

                                            End If

                                        Next

                                    End If

                                    'END GOODSRCV_PA loop'

                                    SQLString = "Select ISNULL(SUM(GRA_IPR_LOT_OUTSTANDING),0) AS GRA_IPR_LOT_OUTSTANDING from WMS_GOODSRCV_PA where GR_CODE='" + row("GR_CODE") + "'"
                                    Dim dtTtlGRALOTQTY As New DataTable
                                    dtTtlGRALOTQTY = gDB.getDataTable(SQLString, gConn, transaction)

                                    If dtTtlGRALOTQTY.Rows(0)("GRA_IPR_LOT_OUTSTANDING").ToString.Trim = 0 Then
                                        SQLString = "UPDATE WMS_GOODSRCV SET EBS_UPDATE_STATUS ='Y' WHERE GR_CODE='" & row("GR_CODE").ToString.Trim & "' and GR_STATUS ='POSTED'"
                                        gDB.amendData(SQLString, gConn, transaction)
                                    End If

                                End If

                            Else
                                lblMSG.Text = "IPR GRN already uploaded to EBS"
                            End If

                        End If

                    End If

                Next

                'Delete Temp action records
                SQLString = "delete  from WMS_EBS_TRANS_ITX_ACTION  where [STATUS] in ('TEMP','ERROR') and DOC_TYPE in ('IPR','IPR_ADJ') " &
                            " and (select count(1) from WMS_EBS_TRANS_ITX_IPR_GRN where TRANSACTION_ID=WMS_EBS_TRANS_ITX_ACTION.TRANSACTION_ID)=0 " &
                            " and (select count(1) from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT where TRANSACTION_ID=WMS_EBS_TRANS_ITX_ACTION.TRANSACTION_ID)=0"
                gDB.amendData(SQLString, gConn, transaction)

                'For INTERFACE LOG
                If IPRInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "IPR_GRN_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_IPR_GRN_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_IPR_GRN"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='IPR'"
                    Dim IPRITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If IPRITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='IPR' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "IPR", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            'If lblMSG.Text = "" Then
            Dim strAlert As String = "<p>IPR GRN Row Inserted : " & insertCount.ToString.Trim & "</p>"
            strAlert &= "<p>IPR GRN Row Updated : " & updateCount.ToString.Trim & "</p>"
            strAlert &= "<p>STOCK Relocation Row Inserted : " & insertCountADJ.ToString.Trim & "</p>"
            'strAlert &= "<p>IPR GRN data has been successfully imported!!</p>"
            lblMSG.Text = strAlert
            'Else
            '    lblMSG.Text = "PO_GRN data has been already updated"
            'End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import STOCK_RETURN
    Public Sub ImportSR_GRNData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim insertCountD As Int16 = 0
        Dim updateCountD As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select a.*,b.RT_REF_NO,b.RT_STATUS,b.RT_DATE FROM WMS_STOCK_RETURN_D a Inner join WMS_STOCK_RETURN b on a.RT_CODE=b.RT_CODE WHERE b.RT_STATUS ='POSTED' and a.EBS_UPDATE_STATUS='N'"
            Dim codtGRCV As New DataTable
            codtGRCV = gDB.getDataTable(SQLString, gConn, transaction)

            If codtGRCV IsNot Nothing AndAlso codtGRCV.Rows.Count > 0 Then
                Dim SRInsert As Int16 = 0

                For Each rowGRCV As DataRow In codtGRCV.Rows
                    If (rowGRCV("RTD_REF_NO").ToString.Trim) <> "" Then
                        'FOR BATCH_NO'
                        Dim dtBatchNo As New DataTable
                        SQLString = "Select REPLACE('WMSSR'+Convert(varchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,6,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSSR1','WMSSR') as BATCH_NO From WMS_EBS_TRANS_ITX_SO_RETURN"
                        dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                        'FOR EBS_WMS_TRANS_ITX_ACTION'
                        Dim dtACTION As New DataTable
                        SQLString = "Select TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,ACTION,STATUS,CREATION_DATE,LAST_UPDATE_DATE from EBS_WMS_TRANS_ITX_ACTION WHERE TRANSACTION_ID='" & rowGRCV("RTD_REF_NO").ToString.Trim() & "'"
                        dtACTION = gDB.getDataTable(SQLString, gConn, transaction)

                        If dtACTION IsNot Nothing AndAlso dtACTION.Rows.Count > 0 Then

                            For Each rowACTION As DataRow In dtACTION.Rows
                                'for TRANSACTION_ID as use seq_no'
                                Dim dtSeq As New DataTable
                                SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                                dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

                                'FOR ACTION'
                                'INSERT
                                SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,EBS_TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,STATUS,BATCH_NO,CREATION_DATE,LAST_UPDATE_DATE) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@EBS_TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@DOC_ID,@STATUS,@BATCH_NO,@CREATION_DATE,@LAST_UPDATE_DATE)"
                                cmd = New SqlCommand(SQLString, gConn, transaction)
                                cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", rowGRCV("RTD_REF_NO").ToString.Trim())
                                cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@STATUS", "NEW")
                                cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                cmd.CommandType = System.Data.CommandType.Text
                                cmd.ExecuteScalar()

                                '      SQLString = "Select ACTION,DOC_TYPE,EBS_TRANSACTION_ID,TRANSACTION_ID,BATCH_NO from WMS_EBS_TRANS_ITX_ACTION " &
                                '"WHERE STATUS <> 'COMPLETED' and STATUS <> 'ERROR' and DOC_TYPE='SO_RETURN' and EBS_TRANSACTION_ID='" & rowGRCV("RTD_REF_NO").ToString.Trim() & "'"
                                '      Dim codtl As New DataTable
                                '      codtl = gDB.getDataTable(SQLString, gConn, transaction)

                                '      If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then

                                '          If (codtl.Rows(0)("ACTION").ToString.Trim) = "NEW" And (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "SO_RETURN" Then
                                'for seq_no'
                                Dim dtTO As New DataTable
                                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_SO_RETURN"
                                dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                                Dim LOC As String
                                Dim RTDlocation As String = ""
                                LOC = rowGRCV("RTD_LOC").ToString.Trim
                                If LOC <> "" Then
                                    RTDlocation = LOC.Substring(0, 2) + "." + LOC.Substring(2, 2) + "." + LOC.Substring(4, 2) + "." + LOC.Substring(6, 2)
                                Else
                                    SQLString = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists' WHERE EBS_TRANSACTION_ID='" & (rowGRCV("RTD_REF_NO").ToString.Trim()) & "'"
                                    gDB.amendData(SQLString, gConn, transaction)
                                End If

                                SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_RETURN " &
                    "WHERE TRANSACTION_ID ='" & rowGRCV("RTD_REF_NO").ToString.Trim & "' and SO_HEADER_ID='" & rowGRCV("SAP_MAT_DOC_ITEM").ToString.Trim & "' and LINE_ID='" & rowGRCV("RTD_PALLET_NO").ToString.Trim & "'"
                                Dim dtROCode As New DataTable
                                dtROCode = gDB.getDataTable(SQLString, gConn, transaction)

                                'INSERT
                                SQLString = "Insert into WMS_EBS_TRANS_ITX_SO_RETURN(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,SO_HEADER_ID,SO_NO,IO_ID,COMPANY_ID,LINE_ID,LINE_NUMBER,ITEM_ID,ITEM_NUMBER,SUBINVENTORY_CODE,CUSTOMER_ID,REF_CUSTOMER_ID,UOM_CODE,QUANTITY,DEST_SUBINVENTORY_CODE,DEST_LOT_NUMBER,DEST_LOT_EXPIRY_DATE,DEST_LOCATOR,DEST_REC_UOM_CODE,DEST_REC_QUANTITY,CREATION_DATE,CREATION_BY,LAST_UPDATE_DATE,EBS_TRANSACTION_ID) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@SO_HEADER_ID,@SO_NO,@IO_ID,@COMPANY_ID,@LINE_ID,@LINE_NUMBER,@ITEM_ID,@ITEM_NUMBER,@SUBINVENTORY_CODE,@CUSTOMER_ID,@REF_CUSTOMER_ID,@UOM_CODE,@QUANTITY,@DEST_SUBINVENTORY_CODE,@DEST_LOT_NUMBER,@DEST_LOT_EXPIRY_DATE,@DEST_LOCATOR,@DEST_REC_UOM_CODE,@DEST_REC_QUANTITY,@CREATION_DATE,@CREATION_BY,@LAST_UPDATE_DATE,@EBS_TRANSACTION_ID)"
                                cmd = New SqlCommand(SQLString, gConn, transaction)
                                cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", rowGRCV("RT_DATE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@SO_HEADER_ID", dtROCode.Rows(0)("SO_HEADER_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@SO_NO", dtROCode.Rows(0)("SO_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@IO_ID", dtROCode.Rows(0)("IO_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@COMPANY_ID", dtROCode.Rows(0)("COMPANY_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@LINE_ID", dtROCode.Rows(0)("LINE_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@LINE_NUMBER", dtROCode.Rows(0)("LINE_NUMBER").ToString.Trim)
                                cmd.Parameters.AddWithValue("@ITEM_ID", dtROCode.Rows(0)("ITEM_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@ITEM_NUMBER", dtROCode.Rows(0)("ITEM_NUMBER").ToString.Trim)
                                cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", dtROCode.Rows(0)("SUBINVENTORY_CODE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@CUSTOMER_ID", dtROCode.Rows(0)("CUSTOMER_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@REF_CUSTOMER_ID", dtROCode.Rows(0)("REF_CUSTOMER_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@UOM_CODE", dtROCode.Rows(0)("UOM_CODE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@QUANTITY", dtROCode.Rows(0)("QUANTITY").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DEST_LOT_NUMBER", rowGRCV("RTD_BATCH_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DEST_LOT_EXPIRY_DATE", rowGRCV("RTD_EXPIRY_DATE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DEST_SUBINVENTORY_CODE", rowGRCV("RTD_WH").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DEST_LOCATOR", RTDlocation)
                                cmd.Parameters.AddWithValue("@DEST_REC_UOM_CODE", rowGRCV("RTD_UOM2").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DEST_REC_QUANTITY", rowGRCV("RTD_QTY2").ToString.Trim)
                                cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", rowGRCV("RTD_REF_NO").ToString.Trim())
                                cmd.Parameters.AddWithValue("@CREATION_DATE", rowGRCV("SYS_LUD").ToString.Trim)
                                cmd.Parameters.AddWithValue("@CREATION_BY", rowGRCV("SYS_CB").ToString.Trim)
                                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", rowGRCV("SYS_LUD").ToString.Trim)
                                cmd.CommandType = System.Data.CommandType.Text
                                cmd.ExecuteScalar()
                                insertCount = insertCount + 1
                                SRInsert = 1

                                SQLString = "UPDATE C SET c.EBS_UPDATE_STATUS ='Y', O.EBS_UPDATE_DATE = GetDate() FROM WMS_STOCK_RETURN_D C Inner JOIN WMS_STOCK_RETURN O ON C.RT_CODE = O.RT_CODE WHERE c.RTD_REF_NO='" & rowGRCV("RTD_REF_NO").ToString.Trim & "' and o.RT_STATUS ='POSTED'"
                                gDB.amendData(SQLString, gConn, transaction)



                                '    End If

                                'Else
                                '    lblMSG.Text = "SO RETURN already uploaded to EBS"
                                'End If

                            Next



                        End If

                    End If

                Next

                'For INTERFACE LOG
                If SRInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "SR_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_SR_EXP_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_SO_RETURN"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                            "WHERE ACTION <> 'CLOSED' and DOC_TYPE='SO_RETURN'"
                    Dim SRITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If SRITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='SO_RETURN' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "SO_RETURN", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>SR Inserted Row : " & insertCount.ToString.Trim & "</p>"
                'strAlert &= "<p>SR Updated Row : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>SR data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "Stock Return data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

    End Sub

    'Import SO_GRN
    Public Sub ImportSO_GRNData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = " Declare @val Varchar(MAX); " &
"select @val = COALESCE(@val + + ',', '')+ DOD_CO_CODE  from [dbo].[WMS_DELV_ORDER_D] dod  " &
"inner Join [dbo].[WMS_DELV_ORDER]  do on dod.DO_CODE=do.DO_CODE " &
"where do.DO_STATUS='POSTED' and   DO_POSTED_DATE>GETDATE()-5 " &
"Select * From WMS_CUST_ORDER Where CO_STATUS !='NEW' And EBS_UPDATE_STATUS='N' " &
"And CO_CODE In  (Select Items From dbo.Split(@val,',')) "
            Dim codtCO As New DataTable
            codtCO = gDB.getDataTable(SQLString, gConn, transaction)
            If codtCO IsNot Nothing AndAlso codtCO.Rows.Count > 0 Then
                Dim SOInsert As Int16 = 0
                For Each rowCOdtl As DataRow In codtCO.Rows
                    'FOR BATCH_NO'
                    Dim BatchRefNo As String
                    BatchRefNo = ""
                    Dim dtBatchNo As New DataTable
                    SQLString = "Select REPLACE('WMSSO'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll(Max(Substring(BATCH_NO,6,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSSO1','WMSSO') as BATCH_NO From WMS_EBS_TRANS_ITX_SO"
                    dtBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim dtACTION As New DataTable
                    SQLString = "Select TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,ACTION,STATUS,CREATION_DATE,LAST_UPDATE_DATE from EBS_WMS_TRANS_ITX_ACTION WHERE TRANSACTION_ID='" & rowCOdtl("CO_CUS_REF_NO").ToString.Trim & "'"
                    dtACTION = gDB.getDataTable(SQLString)

                    If dtACTION IsNot Nothing AndAlso dtACTION.Rows.Count > 0 Then
                        For Each rowACTION As DataRow In dtACTION.Rows
                            'for TRANSACTION_ID as use seq_no'
                            Dim dtSeq As New DataTable
                            SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
                            dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

                            'FOR ACTION'
                            SQLString = "Select * from WMS_EBS_TRANS_ITX_ACTION where EBS_TRANSACTION_ID ='" & rowCOdtl("CO_CUS_REF_NO").ToString.Trim() & "'"
                            Dim dtTRCode As New DataTable
                            dtTRCode = gDB.getDataTable(SQLString, gConn, transaction)
                            If dtTRCode Is Nothing Or dtTRCode.Rows.Count <= 0 Then
                                'INSERT
                                SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,EBS_TRANSACTION_ID,IO_ID,COMPANY_ID,DOC_TYPE,DOC_ID,STATUS,BATCH_NO,CREATION_DATE,LAST_UPDATE_DATE) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@EBS_TRANSACTION_ID,@IO_ID,@COMPANY_ID,@DOC_TYPE,@DOC_ID,@STATUS,@BATCH_NO,@CREATION_DATE,@LAST_UPDATE_DATE)"
                                cmd = New SqlCommand(SQLString, gConn, transaction)
                                cmd.Parameters.AddWithValue("@SOURCE", "WMS")
                                cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", rowCOdtl("CO_CUS_REF_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@STATUS", "NEW")
                                cmd.Parameters.AddWithValue("@BATCH_NO", dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                                cmd.CommandType = System.Data.CommandType.Text
                                cmd.ExecuteScalar()
                                BatchRefNo = dtBatchNo.Rows(0)("BATCH_NO").ToString.Trim
                            Else
                                'UPDATE
                                sbCmdText = New StringBuilder()
                                sbCmdText.Append("Update WMS_EBS_TRANS_ITX_ACTION Set ")
                                sbCmdText.Append("ACTION = @ACTION,")
                                sbCmdText.Append("IO_ID = @IO_ID,")
                                sbCmdText.Append("COMPANY_ID = @COMPANY_ID,")
                                sbCmdText.Append("DOC_TYPE = @DOC_TYPE,")
                                sbCmdText.Append("DOC_ID = @DOC_ID,")
                                sbCmdText.Append("LAST_UPDATE_DATE = @LAST_UPDATE_DATE")
                                sbCmdText.Append(" Where EBS_TRANSACTION_ID = @EBS_TRANSACTION_ID")
                                cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                                cmd.Parameters.AddWithValue("@ACTION", rowACTION("ACTION").ToString.Trim)
                                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", rowCOdtl("CO_CUS_REF_NO").ToString.Trim)
                                cmd.Parameters.AddWithValue("@IO_ID", rowACTION("IO_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@COMPANY_ID", rowACTION("COMPANY_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_TYPE", rowACTION("DOC_TYPE").ToString.Trim)
                                cmd.Parameters.AddWithValue("@DOC_ID", rowACTION("DOC_ID").ToString.Trim)
                                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", System.DateTime.Now)
                                cmd.CommandType = System.Data.CommandType.Text
                                cmd.ExecuteScalar()
                            End If

                        Next

                        SQLString = "Select ACTION,DOC_TYPE,EBS_TRANSACTION_ID,TRANSACTION_ID,BATCH_NO from WMS_EBS_TRANS_ITX_ACTION " &
"WHERE STATUS <> 'COMPLETED' and STATUS <> 'ERROR' and DOC_TYPE='SO' and EBS_TRANSACTION_ID='" & rowCOdtl("CO_CUS_REF_NO").ToString.Trim & "' order by TRANSACTION_ID desc"

                        Dim codtl As New DataTable
                        codtl = gDB.getDataTable(SQLString, gConn, transaction)

                        If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then

                            If (codtl.Rows(0)("ACTION").ToString.Trim) = "NEW" And (codtl.Rows(0)("DOC_TYPE").ToString.Trim) = "SO" Then

                                SQLString = "Select * from WMS_CUST_ORDER_D Where CO_CODE='" & rowCOdtl("CO_CODE").ToString.Trim & "'"
                                Dim codtCODtl As New DataTable
                                codtCODtl = gDB.getDataTable(SQLString, gConn, transaction)
                                If codtCODtl IsNot Nothing AndAlso codtCODtl.Rows.Count > 0 Then
                                    For Each rowCOD As DataRow In codtCODtl.Rows
                                        SQLString = "Select * From WMS_DELV_ORDER Where DO_STATUS='POSTED' and ','+Replace(DO_CO_CODE,' ','')+',' like '%," & rowCOD("CO_CODE").ToString.Trim & ",%' " 'dbo.fun_QueryCSVColumn (Replace(DO_CO_CODE,' ',''),',','" & rowCOD("CO_CODE").ToString.Trim & "','exact contains')='1'"
                                        Dim codtDO As New DataTable
                                        codtDO = gDB.getDataTable(SQLString, gConn, transaction)
                                        If codtDO IsNot Nothing AndAlso codtDO.Rows.Count > 0 Then
                                            Dim CO_QTY As Double = Convert.ToDouble(rowCOD("COD_QTY").ToString)
                                            For Each rowdtl As DataRow In codtDO.Rows

                                                SQLString = "SELECT a.STORER_CODE,a.IMP_CODE,a.PLD_VEND_SEG,a.PLD_SERIAL_NO,a.PLD_PALLET_NO,a.do_code,a.pld_seq,a.pld_item_qty,a.pld_bal_qty,a.PLD_BATCH_NO,a.PLD_LOC,a.PLD_WH,a.PLD_DO_QTY,a.SYS_CD,a.SYS_CB,a.SYS_LUD, a.DOD_SEQ, b.DOD_QTY, b.DOD_CO_CODE,PLD_ITEM_QTY-IsNull(PLD_BAL_QTY,0) as Bal_Qty FROM dbo.WMS_DO_PICKLIST_D a, WMS_DELV_ORDER_D b where a.DO_CODE = b.do_code and a.DOD_SEQ = b.DOD_SEQ and a.do_code='" & rowdtl("DO_CODE").ToString.Trim & "' and PLD_ITEM_NO= '" & rowCOD("COD_ITM_CODE").ToString.Trim & "' and PLD_ITEM_QTY-IsNull(PLD_BAL_QTY,0) > 0 Order by Bal_Qty desc"
                                                Dim codtGRDRP As New DataTable
                                                codtGRDRP = gDB.getDataTable(SQLString, gConn, transaction)
                                                If codtGRDRP IsNot Nothing AndAlso codtGRDRP.Rows.Count > 0 Then

                                                    For Each rowPickdtl As DataRow In codtGRDRP.Rows
                                                        If CO_QTY > 0 Then
                                                            Dim Dest_QTY As Double = 0
                                                            If Convert.ToDouble(rowPickdtl("BAL_QTY").ToString) >= CO_QTY Then
                                                                Dest_QTY = CO_QTY
                                                            Else
                                                                Dest_QTY = Convert.ToDouble(rowPickdtl("BAL_QTY").ToString)
                                                            End If

                                                            'for seq_no'
                                                            Dim dtTO As New DataTable
                                                            SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_SO"
                                                            dtTO = gDB.getDataTable(SQLString, gConn, transaction)

                                                            'for CUSTOMER CODE'
                                                            SQLString = "Select TOP(1) ACCOUNT_NUMBER from EBS_WMS_TRANS_ITX_SO_HEADER where TRANSACTION_ID='" & rowCOdtl("CO_CUS_REF_NO").ToString.Trim & "' and SO_HEADER_ID='" & rowCOdtl("CO_PROJECT_NO").ToString.Trim & "'"
                                                            Dim dtCUSCODE As DataTable
                                                            dtCUSCODE = gDB.getDataTable(SQLString, gConn, transaction)

                                                            'for DEST_LOCATOR
                                                            SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + rowPickdtl("PLD_WH").ToString.Trim() + rowPickdtl("PLD_LOC").ToString.Trim() + "' and IO_ID='" + rowPickdtl("STORER_CODE").ToString.Trim() + "'"
                                                            Dim dtLOC As New DataTable
                                                            dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
                                                            Dim LocationId As String = "0"
                                                            If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                                                                LocationId = dtLOC.Rows(0)("LOCID").ToString()
                                                            Else
                                                                SQLString = "UPDATE WMS_EBS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='Location does not exists' WHERE EBS_TRANSACTION_ID='" & (codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim) & "'"
                                                                gDB.amendData(SQLString, gConn, transaction)
                                                            End If

                                                            'for DEST_LOCATION_ID'
                                                            Dim LOC As String
                                                            Dim PLDlocation As String
                                                            LOC = rowPickdtl("PLD_LOC").ToString.Trim()
                                                            If LOC <> "" Then
                                                                PLDlocation = LOC.Substring(0, 2) + "." + LOC.Substring(2, 2) + "." + LOC.Substring(4, 2) + "." + LOC.Substring(6, 2)
                                                            Else
                                                                PLDlocation = ""
                                                            End If

                                                            SQLString = "Select * from EBS_WMS_TRANS_ITX_SO_DETAIL where TRANSACTION_ID='" & rowCOD("COD_JOB_NO").ToString.Trim & "' and SO_HEADER_ID='" & rowCOD("COD_REF_NO").ToString.Trim & "' and SO_LINE_ID='" & rowCOD("COD_PACKING").ToString.Trim & "'"
                                                            Dim codt As New DataTable
                                                            codt = gDB.getDataTable(SQLString, gConn, transaction)

                                                            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                                                                SQLString = "Select * from WMS_EBS_TRANS_ITX_SO where TRANSACTION_ID ='" & codt.Rows(0)("TRANSACTION_ID").ToString.Trim & "' and SO_HEADER_ID='" & codt.Rows(0)("SO_HEADER_ID").ToString.Trim & "' and SO_LINE_ID='" & codt.Rows(0)("SO_LINE_ID").ToString.Trim & "'"
                                                                Dim codtSO As New DataTable
                                                                codtSO = gDB.getDataTable(SQLString, gConn, transaction)

                                                                If codtSO Is Nothing Or codtSO.Rows.Count <= 0 Then
                                                                    'INSERT
                                                                    SQLString = "Insert into WMS_EBS_TRANS_ITX_SO(SEQ_NO,BATCH_NO,TRANSACTION_ID,ACCOUNT_NUMBER,TRANSACTION_DATE,SO_HEADER_ID,SO_LINE_ID,IO_ID,DELIVERY_ID,DELIVERY_DETAIL_ID,MOVE_ORDER_HEADER_ID,MOVE_ORDER_LINE_ID,ITEM_ID,ITEM_NUMBER,ORDERED_PRIMARY_QUANTITY,ORDERED_PRIMARY_UOM,DEST_LOT_NUMBER,DEST_LOCATION_ID,DEST_LOCATOR,DEST_QUANTITY,DEST_SUBINVENTORY_CODE,CREATION_DATE,CREATION_BY,LAST_UPDATE_DATE,EBS_TRANSACTION_ID) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@ACCOUNT_NUMBER,@TRANSACTION_DATE,@SO_HEADER_ID,@SO_LINE_ID,@IO_ID,@DELIVERY_ID,@DELIVERY_DETAIL_ID,@MOVE_ORDER_HEADER_ID,@MOVE_ORDER_LINE_ID,@ITEM_ID,@ITEM_NUMBER,@ORDERED_PRIMARY_QUANTITY,@ORDERED_PRIMARY_UOM,@DEST_LOT_NUMBER,@DEST_LOCATION_ID,@DEST_LOCATOR,@DEST_QUANTITY,@DEST_SUBINVENTORY_CODE,@CREATION_DATE,@CREATION_BY,@LAST_UPDATE_DATE,@EBS_TRANSACTION_ID)"
                                                                    'SHIP_TO_CUSTOMER_ID,@SHIP_TO_CUSTOMER_ID
                                                                    cmd = New SqlCommand(SQLString, gConn, transaction)
                                                                    cmd.Parameters.AddWithValue("@SEQ_NO", dtTO.Rows(0)("SEQ_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@BATCH_NO", codtl.Rows(0)("BATCH_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_ID", codtl.Rows(0)("TRANSACTION_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ACCOUNT_NUMBER", dtCUSCODE.Rows(0)("ACCOUNT_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@TRANSACTION_DATE", rowPickdtl("SYS_LUD").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@SO_HEADER_ID", codt.Rows(0)("SO_HEADER_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@SO_LINE_ID", codt.Rows(0)("SO_LINE_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@IO_ID", codt.Rows(0)("IO_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@DELIVERY_ID", codt.Rows(0)("DELIVERY_ID").ToString.Trim())
                                                                    cmd.Parameters.AddWithValue("@DELIVERY_DETAIL_ID", codt.Rows(0)("DELIVERY_DETAIL_ID").ToString.Trim())
                                                                    cmd.Parameters.AddWithValue("@MOVE_ORDER_HEADER_ID", codt.Rows(0)("MOVE_ORDER_HEADER_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@MOVE_ORDER_LINE_ID", codt.Rows(0)("MOVE_ORDER_LINE_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ITEM_ID", codt.Rows(0)("ITEM_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ITEM_NUMBER", codt.Rows(0)("ITEM_NUMBER").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ORDERED_PRIMARY_QUANTITY", codt.Rows(0)("ORDERED_PRIMARY_QUANTITY").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@ORDERED_PRIMARY_UOM", codt.Rows(0)("ORDERED_PRIMARY_UOM").ToString.Trim)
                                                                    'cmd.Parameters.AddWithValue("@SHIP_TO_CUSTOMER_ID", codt.Rows(0)("UOM_CODE").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@DEST_LOT_NUMBER", rowPickdtl("PLD_BATCH_NO").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@DEST_LOCATOR", PLDlocation)
                                                                    cmd.Parameters.AddWithValue("@DEST_SUBINVENTORY_CODE", rowPickdtl("PLD_WH").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@DEST_LOCATION_ID", LocationId)
                                                                    cmd.Parameters.AddWithValue("@DEST_QUANTITY", Dest_QTY)
                                                                    cmd.Parameters.AddWithValue("@EBS_TRANSACTION_ID", codtl.Rows(0)("EBS_TRANSACTION_ID").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@CREATION_DATE", rowPickdtl("SYS_LUD").ToString.Trim)
                                                                    cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_nickname"))
                                                                    cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", rowPickdtl("SYS_LUD").ToString.Trim)
                                                                    cmd.CommandType = System.Data.CommandType.Text
                                                                    cmd.ExecuteScalar()
                                                                    insertCount = insertCount + 1
                                                                    SOInsert = 1

                                                                    SQLString = "UPDATE WMS_DO_PICKLIST_D SET PLD_BAL_QTY =ISNULL(PLD_BAL_QTY,0) + " & Math.Round(Dest_QTY, 2) & " WHERE DO_CODE='" & rowdtl("DO_CODE").ToString.Trim & "' And PLD_SEQ = '" + rowPickdtl("PLD_SEQ") + "'"
                                                                    gDB.amendData(SQLString, gConn, transaction)
                                                                    CO_QTY = CO_QTY - Dest_QTY

                                                                    Dim Remark As String = "CO_CODE: " + rowCOD("CO_CODE").ToString.Trim + ", DO_CODE: " + rowdtl("DO_CODE").ToString.Trim + ", SO_HEADER_ID: " + codt.Rows(0)("SO_HEADER_ID").ToString.Trim + ", ITEM_CODE: " + rowCOD("COD_ITM_CODE").ToString.Trim + ", PICKED_QTY: " + rowPickdtl("pld_item_qty").ToString.Trim
                                                                    Dim insertLog As String = "insert into [dbo].[ActionLogs] values ('" + gU.dbEncode(Remark) + "','SO_UPLOAD','WMS_CUST_ORDER',GetDate(),'0')"
                                                                    gDB.amendData(insertLog, gConn, transaction)

                                                                End If

                                                            End If

                                                        End If

                                                    Next

                                                End If

                                            Next

                                        End If

                                    Next

                                    SQLString = "UPDATE WMS_CUST_ORDER SET EBS_UPDATE_STATUS ='Y' WHERE CO_CODE='" & rowCOdtl("CO_CODE").ToString.Trim & "' and CO_STATUS! ='NEW'"
                                    gDB.amendData(SQLString, gConn, transaction)
                                End If

                            End If

                        Else
                            lblMSG.Text = "SO already uploaded to EBS"
                        End If

                    End If
                    SQLString = "update WMS_EBS_TRANS_ITX_ACTION set RELATED_NO = '" & BatchRefNo & "' where DOC_TYPE='LOT_ADJ' and ISNULL(RELATED_NO,'')=''"
                    gDB.amendData(SQLString, gConn, transaction)
                Next

                'For INTERFACE LOG
                If SOInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "SO_EXP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_SO_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If


                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_EBS_TRANS_ITX_SO"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)


                    SQLString = "Select STATUS from WMS_EBS_TRANS_ITX_ACTION " &
                        "WHERE ACTION <> 'CLOSED' and DOC_TYPE='SO'"
                    Dim SOITXACTION As DataTable = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    If SOITXACTION.Rows(0)("STATUS").ToString.Trim = "NEW" Then
                        ITFSTATUS = "SUCCESS"
                    Else
                        ITFSTATUS = "ERROR"
                    End If

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='SO' and ITF_TYPE='E'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "E", "SO", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            'If lblMSG.Text = "" Then
            Dim strAlert As String = "<p>SO Inserted Row : " & insertCount.ToString.Trim & "</p>"
            'strAlert &= "<p>SO Updated Row : " & updateCount.ToString.Trim & "</p>"
            'strAlert &= "<p>SO data has been successfully imported!!</p>"
            lblMSG.Text = strAlert
            'Else
            ' lblMSG.Text = "SO data has been already updated"
            'End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'FOR Interface Log'
    Private Sub InterfaceLog(IMP_BATCH_ID As String, IMP_FILE_NAME As String, ITF_TYPE As String, ITF_IMP_TYPE As String, ITF_START_DT As DateTime, ITF_END_DT As DateTime, ITF_STATUS As String, ITF_ERRCODE As String, ITF_ERRDESC As String, ITF_TOT_RECS As Int32, ITF_TOT_SUCCESS As Int32, ITF_TOT_ERRORED As Int32)
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        'Dim SQLString As String
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        'Try

        '    SQLString = "Insert into WMS_INTF_LOG(IMP_BATCH_ID,IMP_FILE_NAME,ITF_TYPE,ITF_IMP_TYPE,ITF_START_DT,ITF_END_DT,ITF_STATUS,ITF_ERRCODE,ITF_ERRDESC,ITF_TOT_RECS,ITF_TOT_SUCCESS,ITF_TOT_ERRORED) VALUES (@IMP_BATCH_ID,@IMP_FILE_NAME,@ITF_TYPE,@ITF_IMP_TYPE,@ITF_START_DT,@ITF_END_DT,@ITF_STATUS,@ITF_ERRCODE,@ITF_ERRDESC,@ITF_TOT_RECS,@ITF_TOT_SUCCESS,@ITF_TOT_ERRORED)"
        '    cmd = New SqlCommand(SQLString, gConn, transaction)
        '    cmd.Parameters.AddWithValue("@IMP_BATCH_ID", IMP_BATCH_ID)
        '    cmd.Parameters.AddWithValue("@IMP_FILE_NAME", IMP_FILE_NAME)
        '    cmd.Parameters.AddWithValue("@ITF_TYPE", ITF_TYPE)
        '    cmd.Parameters.AddWithValue("@ITF_IMP_TYPE", ITF_IMP_TYPE)
        '    cmd.Parameters.AddWithValue("@ITF_START_DT", ITF_START_DT)
        '    cmd.Parameters.AddWithValue("@ITF_END_DT", ITF_END_DT)
        '    cmd.Parameters.AddWithValue("@ITF_STATUS", ITF_STATUS)
        '    cmd.Parameters.AddWithValue("@ITF_ERRCODE", ITF_ERRCODE)
        '    cmd.Parameters.AddWithValue("@ITF_ERRDESC", ITF_ERRDESC)
        '    cmd.Parameters.AddWithValue("@ITF_TOT_RECS", ITF_TOT_RECS)
        '    cmd.Parameters.AddWithValue("@ITF_TOT_SUCCESS", ITF_TOT_SUCCESS)
        '    cmd.Parameters.AddWithValue("@ITF_TOT_ERRORED", ITF_TOT_ERRORED)
        '    cmd.CommandType = System.Data.CommandType.Text
        '    cmd.ExecuteScalar()

        '    transaction.Commit()

        'Catch ex As Exception
        '    Response.Write(ex.Message)
        '    uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        '    If Not transaction Is Nothing Then
        '        transaction.Rollback()
        '        transaction = Nothing
        '    End If
        'Finally
        '    If gConn IsNot Nothing Then
        '        If gConn.State = ConnectionState.Open Then
        '            gConn.Close()
        '            gConn.Dispose()
        '        End If
        '    End If
        'End Try

    End Sub

    'For get last lot no
    Function getLastLot(ByVal ITEM_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String) As String

        Dim LAST_LOT As String = ""
        Dim SQLString As String = "Select iSnuLL(DEST_LOT_NUMBER,'')LAST_LOT from WMS_EBS_TRANS_ITX_SO so Where so.ITEM_ID = '" + ITEM_CODE + "' and so.ACCOUNT_NUMBER='" + CUS_CODE + "' and so.IO_ID = '" + STORER_CODE + "'"
        Dim dtItem As DataTable = gDB.getDataTable(SQLString)
        If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
            LAST_LOT = dtItem.Rows(0)("LAST_LOT")
        End If
        Return LAST_LOT
    End Function

    Public Sub POSTDATASTKADJ(STKADJCodes As String)
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim updtSql As String
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Dim STORER_CODE As String = ""
        Dim nextNo As String = ""
        Dim WH_CODE As String = ""
        Dim currTrans As String = ""
        Try
            Dim AdjCodes = STKADJCodes.Split(",")
            For Each ADJCode As String In AdjCodes
                SQLString = "select * from WMS_STOCK_ADJUST where AD_CODE='" & ADJCode & "'  "
                Dim dtAdj As DataTable
                dtAdj = gDB.getDataTable(SQLString, gConn, transaction)
                If dtAdj IsNot Nothing AndAlso dtAdj.Rows.Count > 0 Then
                    STORER_CODE = dtAdj.Rows(0).Item("STORER_CODE").ToString
                    nextNo = dtAdj.Rows(0).Item("AD_CODE").ToString
                    WH_CODE = dtAdj.Rows(0).Item("AD_WH").ToString
                    currTrans = dtAdj.Rows(0).Item("AD_REF_NO").ToString

                    SQLString = "SELECT [IMP_CODE],[STORER_CODE],[AD_CODE],[AD_SEQ],[ADD_ITM_CODE],[ADD_PACK_KEY] " &
",[ADD_LOC],[ADD_ORG_QTY],[ADD_REV_QTY],[ADD_VAR_QTY],[ADD_REM],[ADD_PALLET_NO] " &
",[SYS_LUB],[SYS_LUD],[SYS_CD],[SYS_CB],[ADD_BATCH_NO] " &
",[ADD_VND_CODE],convert(varchar(10),[ADD_EXPIRY_DATE],103) ADD_EXPIRY_DATE,convert(varchar(10),[ADD_MANU_DATE],103)ADD_MANU_DATE,[ADD_ORG_QTY2],[ADD_REV_QTY2] " &
",[ADD_VAR_QTY2],[ADD_SERIAL_NO],[DRUM_ID],[DRUM_LEVEL]FROM[dbo].[WMS_STOCK_ADJUST_D] " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and AD_CODE='" & nextNo & "'  "
                    Dim dtRODCode As DataTable
                    dtRODCode = gDB.getDataTable(SQLString, gConn, transaction)

                    If dtRODCode IsNot Nothing AndAlso dtRODCode.Rows.Count > 0 Then
                        For Each rows As DataRow In dtRODCode.Rows
                            'FOR POSTING'
                            If gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "") <> "000" AndAlso gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "") <> "" Then
                                WH_CODE = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                            End If
                            st.STORER_CODE = STORER_CODE
                            st.ITM_CODE = gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "SADJ"
                            st.IO_DOC_ID = nextNo
                            st.IO_CBM = 0
                            st.IO_KG = 0
                            st.IO_WH = WH_CODE
                            st.PALLET_NO = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                            st.IO_LOC = gU.decodeNull(rows.Item("add_loc").ToString.Trim, "")
                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, "")
                            st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, "")
                            st.IO_MANU_DATE = gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "")

                            Dim txQty As Double = 0

                            If gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") >
                                gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                                txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0")

                                st.IO_QTY = txQty
                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                                st.UpdateStockTrans("IN", gConn, transaction)
                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                    st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                    st.IOS_QTY2 = 1
                                    st.UpdateStockSerialTrans("IN", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                End If

                            ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") <
                                gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                                txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                                st.IO_QTY = txQty
                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                                st.UpdateStockTrans("OUT", gConn, transaction)
                                st.UpdateStockBalTrans("OUT", gConn, transaction)

                                If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                    st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                    st.IOS_QTY2 = 1
                                    st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                End If

                            ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") >
                                gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                                txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0")

                                Dim drumDt As DataTable
                                SQLString = "SELECT ILBS_DRUM_ID, ILBS_DRUM_LEVEL FROM WMS_ITEM_LOC_BAL_S WHERE " &
                                     "IMP_CODE = 'WMS' " &
                                            "AND STORER_CODE = '" & STORER_CODE & "' " &
                                            "AND ITM_CODE = '" & gU.decodeNull(rows.Item("ADD_ITM_CODE").ToString, "") & "' " &
                                            "AND PACK_KEY = '1' " &
                                            "AND ILBS_SERIAL_NO = '" & gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") & "' AND ILBS_QTY2 > 0 "
                                REM **********************

                                drumDt = gDB.getDataTable(SQLString, gConn, transaction)

                                If drumDt.Rows.Count > 0 Then
                                    st.IOS_DRUM_ID = drumDt.Rows(0).Item("ILBS_DRUM_ID").ToString
                                    st.IOS_DRUM_LEVEL = drumDt.Rows(0).Item("ILBS_DRUM_LEVEL")
                                End If

                                st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                st.IOS_QTY2 = txQty
                                st.IO_QTY = 1
                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                                st.UpdateStockTrans("IN", gConn, transaction)
                                st.UpdateStockBalTrans("IN", gConn, transaction)
                                If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                    st.IOS_SL = "Y"
                                    st.UpdateStockSerialTrans("IN", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                End If

                            ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") <
                                gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                                txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0")
                                st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                st.IOS_QTY2 = txQty
                                st.IO_QTY = 1
                                Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                                st.UpdateStockTrans("OUT", gConn, transaction)
                                st.UpdateStockBalTrans("OUT", gConn, transaction)
                                If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                    st.IOS_SL = "Y"
                                    st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                End If

                            ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") =
                                gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                                'If transaction IsNot Nothing Then
                                '    transaction.Rollback()
                                'End If
                                Exit Sub
                            End If
                        Next

                        updtSql = "update wms_stock_adjust " &
                                    "set ad_status = 'POSTED' " &
                                    "where imp_code = 'WMS' " &
                                    "and storer_code = '" & STORER_CODE & "' " &
                                    "and ad_code = '" & nextNo & "' "

                        gDB.amendData(updtSql, gConn, transaction)
                    Else
                        'updtSql = "update EBS_WMS_TRANS_ITX_ACTION set [STATUS]='ERROR' , [ERROR_MESSAGE]='-ve stock found' where TRANSACTION_ID='" & TRANSACTION_ID & "'"
                        'gDB.amendData(updtSql, gConn, transaction)

                        updtSql = "update WMS_STOCK_ADJUST set AD_REM='-ve stock found' where AD_CODE='" & nextNo & "'"
                        gDB.amendData(updtSql, gConn, transaction)

                    End If
                End If
            Next
            currTrans = ""
            nextNo = ""
            transaction.Commit()
        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
            updtSql = "UPDATE EBS_WMS_TRANS_ITX_ACTION SET STATUS ='ERROR',ERROR_MESSAGE='" + ex.Message.Replace("'", "") + ", AD_CODE = " + nextNo + "',LAST_UPDATE_DATE='" + System.DateTime.Now + "' WHERE TRANSACTION_ID='" & currTrans & "'"
            gDB.amendData(updtSql)
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    Public Sub LotAllocationCO(transaction As SqlTransaction, gConn As SqlConnection, BatchNo As String)
        Dim SQLString As String
        Dim updateSQL As String = ""
        Dim nextNo As String
        Dim coHdrDATES As DataTable
        Dim coHdrGrp As DataTable
        Dim coDtl As DataTable
        Dim Qty As Double = 0
        Dim SeqNo As Int16 = 0
        Dim OrderCount = 0
        Dim ruleWareList = New List(Of String)
        ruleWareList.Add("FG")
        ruleWareList.Add("VIP")
        Try
            SQLString = "Select Distinct IMP_CODE,STORER_CODE,CO_DATE,ROUTE_ID FROM WMS_CUST_ORDER SS Where ss.CO_STATUS='NEW' and ss.CO_EDI_SIR_NO = '" & BatchNo & "' "
            coHdrDATES = gDB.getDataTable(SQLString, gConn, transaction)

            'And STORER_CODE='" + STORER_CODE + "'

            If coHdrDATES IsNot Nothing AndAlso coHdrDATES.Rows.Count > 0 Then
                For Each dtRow As DataRow In coHdrDATES.Rows
                    Dim STORER_CODE = dtRow("STORER_CODE")
                    Dim IMP_CODE = dtRow("IMP_CODE")
                    SQLString = "SELECT Distinct SS.ROUTE_ID,STUFF((SELECT ', ' + US.CO_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID  and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & dtRow("STORER_CODE") & "' and US.CO_EDI_SIR_NO = '" & BatchNo & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CO_CODES],STUFF((SELECT ', ' + US.CUS_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and US.CO_EDI_SIR_NO = '" & BatchNo & "'  and ss.STORER_CODE = '" & dtRow("STORER_CODE") & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CUS_CODES],STUFF((SELECT ', ' + US.CUS_NAME FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and US.CO_EDI_SIR_NO = '" & BatchNo & "'  and ss.STORER_CODE = '" & dtRow("STORER_CODE") & "' and ss.IMP_CODE = '" & dtRow("IMP_CODE") & "' FOR XML PATH('')), 1, 1, '') [CUS_NAMES] FROM WMS_CUST_ORDER SS Where ss.ROUTE_ID = '" & dtRow("ROUTE_ID") & "' and ss.CO_EDI_SIR_NO = '" & BatchNo & "'  and ss.CO_Date ='" & dtRow("CO_DATE") & "' and ss.CO_STATUS='NEW' and ss.STORER_CODE = '" & dtRow("STORER_CODE") & "' and ss.IMP_CODE = '" & dtRow("IMP_CODE") & "'"
                    coHdrGrp = gDB.getDataTable(SQLString, gConn, transaction)
                    If coHdrGrp IsNot Nothing AndAlso coHdrGrp.Rows.Count > 0 Then
                        coHdrGrp.DefaultView.Sort = "ROUTE_ID asc"
                        coHdrGrp = coHdrGrp.DefaultView.ToTable()
                        For Each dtGrpRow As DataRow In coHdrGrp.Rows
                            nextNo = DB.getDocNo("Do", gConn, transaction)
                            SQLString = "INSERT INTO WMS_DELV_ORDER(IMP_CODE ,STORER_CODE ,DO_CODE,ROUTE_ID ,DO_STATUS ,DO_ISSUED_BY ,DO_CO_CODE ,DO_CUS_REF_NO ,DO_DATE ,CUS_CODE ,CUS_NAME ,DO_POSTED_DATE ,DO_POSTED_BY ,DO_TYPE ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                            SQLString += "VALUES("
                            SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                            SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                            SQLString += "'" & nextNo & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("ROUTE_ID").ToString.Trim) & "',"
                            SQLString += "'DRAFT',"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & "',"
                            SQLString += "'NA',"
                            SQLString += "'" & dtRow("CO_DATE").ToString.Trim & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("CUS_CODES").ToString.Trim) & "',"
                            SQLString += "N'" & gU.dbEncode(dtGrpRow("CUS_NAMES").ToString.Trim) & "'," '
                            SQLString += "GETDATE(),"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "'NA',"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "GETDATE(),"
                            SQLString += "GETDATE(),"
                            SQLString += "'" & Session("usr_id") & "'"
                            SQLString += ")"
                            gDB.amendData(SQLString, gConn, transaction)


                            SQLString = "Select isnull(COD_PCS_UOM,0) COD_PCS_UOM,isnull(COD_TOTPCS,0) COD_TOTPCS,isnull(COD_TOT_WGT,0) COD_TOT_WGT, " &
" isnull(COD_TOT_CBM,0) COD_TOT_CBM,(case when LEN(COD_PALLET_NO)<8 then '19000101' else COD_PALLET_NO end) COD_PALLET_NO,a.*,b.ROUTE_ID,IsNUll(a.COD_WH_CODE,'') as COD_WH,IsNull(c.WH_TYPE,'')WH_TYPE,iSnuLL(b.CUS_CODE,'')CUS_CODE,iSnULL(b.CUS_NAME,'')CUS_NAME from WMS_CUST_ORDER_D a Inner join WMS_CUST_ORDER b On a.CO_CODE = b.CO_CODE Inner join WMS_WAREHOUSE c On a.COD_WH_CODE = c.WH_CODE and a.IMP_CODE=c.IMP_CODE " &
                            "WHERE  " &
                            "b.STORER_CODE = '" & STORER_CODE & "'" &
                            "and b.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' and a.CO_CODE in (" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & ")  Order by b.ROUTE_ID"
                            coDtl = gDB.getDataTable(SQLString, gConn, transaction)
                            If coDtl IsNot Nothing AndAlso coDtl.Rows.Count > 0 Then
                                coDtl.Columns.Add("IsRule")
                                coDtl.Columns.Add("MIN_SELF_LIFE")
                                coDtl.Columns.Add("MIN_PROD_LIFE")
                                coDtl.Columns.Add("MAX_BATCHES")
                                coDtl.Columns.Add("MPL_FLAG")
                                coDtl.Columns.Add("MB_FLAG")
                                coDtl.Columns.Add("MSL_FLAG")
                                coDtl.Columns.Add("LOTS_CANNOT_BE_EARLIER")

                                Dim dtWPickNORULE As New DataTable
                                Dim dtWPick1RULE As New DataTable
                                Dim dtWPickRULE As New DataTable
                                Dim grpQTY As Double = 0
                                Dim grpCus As String = ""
                                Dim custCodes As New List(Of String)
                                Dim custNames As New List(Of String)
                                Dim coCodes As New List(Of String)
                                Dim prevItem As String = coDtl.Rows(0)("COD_ITM_CODE").ToString()
                                Dim RowList As New List(Of DataRow)
                                dtWPickNORULE = coDtl.Clone()
                                dtWPick1RULE = coDtl.Clone()
                                dtWPickRULE = coDtl.Clone()
                                dtWPickNORULE.Rows.Clear()
                                dtWPick1RULE.Rows.Clear()
                                dtWPickRULE.Rows.Clear()

                                For Each row As DataRow In coDtl.Rows
                                    Dim RuleRow = IsRule(gConn, transaction, IMP_CODE, STORER_CODE, row("CUS_CODE").ToString(), row("COD_ITM_CODE").ToString())
                                    If RuleRow IsNot Nothing AndAlso ruleWareList.Contains(row("WH_TYPE").ToString()) = True Then
                                        row("IsRule") = RuleRow("RuleCount")
                                        row("MIN_SELF_LIFE") = RuleRow("MIN_SELF_LIFE")
                                        row("MIN_PROD_LIFE") = RuleRow("MIN_PROD_LIFE")
                                        row("MAX_BATCHES") = RuleRow("MAX_BATCHES")

                                        row("MPL_FLAG") = RuleRow("MPL_FLAG")
                                        row("MSL_FLAG") = RuleRow("MSL_FLAG")
                                        row("MB_FLAG") = RuleRow("MB_FLAG")
                                        row("LOTS_CANNOT_BE_EARLIER") = RuleRow("LOTS_CANNOT_BE_EARLIER")

                                    ElseIf RuleRow Is Nothing OrElse ruleWareList.Contains(row("WH_TYPE").ToString()) = False Then
                                        row("IsRule") = "0"
                                        row("MIN_SELF_LIFE") = "0"
                                        row("MIN_PROD_LIFE") = "0"
                                        row("MAX_BATCHES") = "0"
                                        row("MPL_FLAG") = "0"
                                        row("MSL_FLAG") = "0"
                                        row("MB_FLAG") = "0"
                                        row("LOTS_CANNOT_BE_EARLIER") = "0"
                                    End If
                                Next

                                'No Rule
                                Dim tmpTable = If(coDtl.Select("IsRule = 0 And (WH_TYPE ='FG' OR WH_TYPE ='VIP')").Length > 0, coDtl.Select("IsRule = 0 And (WH_TYPE ='FG' OR WH_TYPE ='VIP')").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' AND COD_WH='" + row("COD_WH").ToString() + "' ")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' AND COD_WH='" + row("COD_WH").ToString() + "' ")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()
                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPickNORULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If


                                    Next
                                End If
                                tmpTable = If(coDtl.Select("IsRule = 0 And WH_TYPE <> 'FG' AND WH_TYPE <> 'VIP'").Length > 0, coDtl.Select("IsRule = 0 And WH_TYPE <> 'FG' AND WH_TYPE <> 'VIP'").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "COD_WH"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    For Each row As DataRow In tmpTable.Rows
                                        Dim rgRow = coDtl.NewRow()
                                        rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                        rgRow("COD_PALLET_NO") = row("COD_PALLET_NO").ToString()
                                        rgRow("COD_CARTON_NO") = row("COD_CARTON_NO").ToString()
                                        rgRow("COD_PACK_KEY") = row("COD_PACK_KEY").ToString()
                                        rgRow("COD_ITM_DESC") = row("COD_ITM_DESC").ToString()
                                        rgRow("COD_UOM") = row("COD_UOM").ToString()
                                        rgRow("COD_PCS_UOM") = If(row("COD_PCS_UOM") IsNot DBNull.Value, row("COD_PCS_UOM").ToString(), "0")
                                        rgRow("COD_TOTPCS") = If(row("COD_TOTPCS") IsNot DBNull.Value, row("COD_TOTPCS").ToString(), "0")
                                        rgRow("COD_TOT_WGT") = If(row("COD_TOT_WGT") IsNot DBNull.Value, row("COD_TOT_WGT").ToString(), "0")
                                        rgRow("COD_TOT_CBM") = If(row("COD_TOT_CBM") IsNot DBNull.Value, row("COD_TOT_CBM").ToString(), "0")
                                        rgRow("COD_QTY") = row("COD_QTY").ToString()
                                        rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                        rgRow("CUS_CODE") = row("CUS_CODE").ToString()
                                        rgRow("CUS_NAME") = row("CUS_NAME").ToString()
                                        rgRow("CO_CODE") = row("CO_CODE").ToString().Trim
                                        rgRow("COD_WH") = row("COD_WH").ToString()
                                        dtWPickNORULE.Rows.Add(rgRow.ItemArray)
                                        RowList.Add(row)
                                    Next
                                End If
                                'No Rule
                                tmpTable = If(coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND MPL_FLAG = 0 AND LOTS_CANNOT_BE_EARLIER = 0").Length > 0, coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND MPL_FLAG = 0 AND LOTS_CANNOT_BE_EARLIER = 0").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "MIN_SELF_LIFE", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows

                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If
                                    Next
                                End If

                                tmpTable = If(coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0 And MPL_FLAG = 0").Length > 0, coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0 And MPL_FLAG = 0").CopyToDataTable(), New DataTable())

                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_PALLET_NO", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows
                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()
                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If
                                    Next
                                End If

                                SeqNo = 1
                                Dim DOD_GRP_SEQ = 1
                                For Each rowD As DataRow In dtWPick1RULE.Rows
                                    Dim LOTNO As String = ""
                                    Dim LOTQTY As Double = 0
                                    Dim chkdList = custRuleChk(gConn, transaction, 1, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)

                                    For Each itemDict In chkdList
                                        SQLString = ""
                                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                        SQLString += "VALUES("
                                        SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                        SQLString += "'" & nextNo & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                        SQLString += "'0',"
                                        SQLString += "'0',"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
                                        SQLString += "0,"
                                        SQLString += "0,"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                        SQLString += "'" & Session("usr_id") & "',"
                                        SQLString += "GETDATE(),"
                                        SQLString += "GETDATE(),"
                                        SQLString += "'" & Session("usr_id") & "'"
                                        SQLString += ")"
                                        SeqNo = SeqNo + 1
                                        gDB.amendData(SQLString, gConn, transaction)
                                    Next
                                    DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                Next
                                tmpTable.Clear()
                                tmpTable = If(coDtl.Select("IsRule > 1").Length > 0, coDtl.Select("IsRule > 1").CopyToDataTable(), New DataTable())

                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    Dim NewTempDt As DataTable
                                    NewTempDt = coDtl.Clone()
                                    NewTempDt.Rows.Clear()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_PALLET_NO", "COD_WH", "MIN_SELF_LIFE", "MIN_PROD_LIFE", "MAX_BATCHES", "MPL_FLAG", "MB_FLAG", "LOTS_CANNOT_BE_EARLIER")
                                    For Each row As DataRow In tmpTableDistinct.Rows

                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "' And MIN_SELF_LIFE ='" + row("MIN_SELF_LIFE").ToString() + "' And MAX_BATCHES ='" + row("MAX_BATCHES").ToString() + "'  And MPL_FLAG ='" + row("MPL_FLAG").ToString() + "' And MB_FLAG ='" + row("MB_FLAG").ToString() + "' And LOTS_CANNOT_BE_EARLIER ='" + row("LOTS_CANNOT_BE_EARLIER").ToString() + "' ")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "' And MIN_SELF_LIFE ='" + row("MIN_SELF_LIFE").ToString() + "' And MAX_BATCHES ='" + row("MAX_BATCHES").ToString() + "' And MPL_FLAG ='" + row("MPL_FLAG").ToString() + "' And MB_FLAG ='" + row("MB_FLAG").ToString() + "' And LOTS_CANNOT_BE_EARLIER ='" + row("LOTS_CANNOT_BE_EARLIER").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            'rgRow.ItemArray = row.ItemArray
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            NewTempDt.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If
                                    Next

                                    For Each rowD As DataRow In NewTempDt.Rows
                                        Dim LOTNO As String = ""
                                        Dim LOTQTY As Double = 0

                                        Dim chkdList = custRuleChk(gConn, transaction, 1, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
                                        For Each itemDict In chkdList
                                            SQLString = ""

                                            SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                            SQLString += "VALUES("
                                            SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                            SQLString += "'" & nextNo & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                            SQLString += "'" & (SeqNo) & "',"
                                            SQLString += "'" & (SeqNo) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                            SQLString += "'0',"
                                            SQLString += "'0',"
                                            SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                            SQLString += "" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & ","
                                            SQLString += "0,"
                                            SQLString += "0,"
                                            SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & ","
                                            SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                            SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                            SQLString += "'" & Session("usr_id") & "',"
                                            SQLString += "GETDATE(),"
                                            SQLString += "GETDATE(),"
                                            SQLString += "'" & Session("usr_id") & "'"
                                            SQLString += ")"
                                            SeqNo = SeqNo + 1
                                            gDB.amendData(SQLString, gConn, transaction)
                                        Next
                                        DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                    Next
                                End If

                                For Each rowD As DataRow In dtWPickNORULE.Rows
                                    Dim LOTNO As String = ""
                                    Dim LOTQTY As Double = 0
                                    Dim chkdList = custRuleChk(gConn, transaction, 0, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
                                    For Each itemDict In chkdList
                                        SQLString = ""
                                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                        SQLString += "VALUES("
                                        SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                        SQLString += "'" & nextNo & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                        SQLString += "'0',"
                                        SQLString += "'0',"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
                                        SQLString += "0,"
                                        SQLString += "0,"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim) & ","
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                        SQLString += "'" & Session("usr_id") & "',"
                                        SQLString += "GETDATE(),"
                                        SQLString += "GETDATE(),"
                                        SQLString += "'" & Session("usr_id") & "'"
                                        SQLString += ")"
                                        SeqNo = SeqNo + 1
                                        gDB.amendData(SQLString, gConn, transaction)
                                    Next
                                    DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                Next

                                SQLString = "Select DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE from WMS_DELV_ORDER_D " &
                                "WHERE  " &
                                "DO_CODE = '" & gU.dbEncode(nextNo) & "'"
                                Dim dtHdr = gDB.getDataTable(SQLString, gConn, transaction)
                                Dim lstName As New List(Of String)
                                Dim lstCode As New List(Of String)
                                Dim lstcoCode As New List(Of String)
                                For Each row As DataRow In dtHdr.Rows
                                    Dim Names = row("DOD_CUS_NAME").ToString().Split(",")
                                    Dim Codes = row("DOD_CUS_CODE").ToString().Split(",")
                                    Dim co_Codes = row("DOD_CO_CODE").ToString().Split(",")
                                    For Each nm As String In Names
                                        lstName.Add(nm)
                                    Next
                                    For Each cd As String In Codes
                                        lstCode.Add(cd)
                                    Next
                                    For Each cocd As String In co_Codes
                                        lstcoCode.Add(cocd.ToString().Trim)
                                    Next
                                Next
                                SQLString = "Update WMS_DELV_ORDER Set CUS_CODE='" + String.Join(",", lstCode.Distinct().ToList()) + "',CUS_NAME=N'" & gU.dbEncode(String.Join(", ", lstName.Distinct().ToList())) & "',DO_CO_CODE=N'" + String.Join(", ", lstcoCode.Distinct().ToList()) + "' WHERE DO_CODE = '" & gU.dbEncode(nextNo) & "'"
                                gDB.amendData(SQLString, gConn, transaction)
                                If lstcoCode IsNot Nothing AndAlso lstcoCode.Count > 0 Then
                                    updateSQL = "Update WMS_CUST_ORDER set CO_STATUS='PICKED'" &
                                        " Where CO_STATUS='NEW' and CO_CODE in (" + String.Join(",", lstcoCode.Distinct().ToList()) + ") "
                                    gDB.amendData(updateSQL, gConn, transaction)
                                End If
                                OrderCount = OrderCount + 1
                            End If
                        Next
                    End If
                Next
            Else
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Throw ex
        End Try
    End Sub

    Private Function IsRule(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal ITEM_CODE As String) As DataRow
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        SQLString = "Select [ITEM_CODE] ,[CUST_CODE] ,isnull(MPL_FLAG,0)MPL_FLAG ,[MIN_PROD_LIFE] ,isnull(MSL_FLAG,0) MSL_FLAG ,[MIN_SELF_LIFE] ,isnull(MB_FLAG,0)MB_FLAG ,ISNULL(MAX_BATCHES,0)MAX_BATCHES ,isnull([LOTS_CANNOT_BE_EARLIER],0) [LOTS_CANNOT_BE_EARLIER] ,[MAX_LOTS], [Series] from WMS_CUSTOMER_RULE where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
        dtRule = gDB.getDataTable(SQLString, conn, transaction)
        If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
            SQLString = "Select isnull(MPL_FLAG,0)MPL_FLAG,MIN_PROD_LIFE,isnull(MSL_FLAG,0) MSL_FLAG,MIN_SELF_LIFE,isnull(MB_FLAG,0)MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,LOTS_CANNOT_BE_EARLIER from WMS_CUSTOMER Where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
            dtRule = gDB.getDataTable(SQLString, conn, transaction)
        End If
        dtRule.Columns.Add("RuleCount")
        If dtRule IsNot Nothing AndAlso dtRule.Rows.Count > 0 Then
            If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 4

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            End If
        Else
            Return Nothing
        End If
        Return dtRule.Rows(0)
    End Function

    Private Function custRuleChk(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal RULE_COUNT As Int16, ByVal WH_CODE As String, ByVal DO_CODE As String, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal ITEM_CODE As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime, ByVal Optional CUS_CODE As String = "", ByVal Optional LAST_LOT As String = "") As List(Of Dictionary(Of String, Object))
        Dim SQLString As String
        Dim dtRule As DataTable
        Dim dtStock As DataTable
        Dim returnList As New List(Of Dictionary(Of String, Object))
        Dim IsFound = 0
        Dim AllotedLotCount = 0
        Try

            SQLString = "Select a.ITM_CODE,a.ITM_SKU_NO,a.ITM_DESC,b.STORER_CODE,b.STO_NAME from WMS_ITEM a Inner Join WMS_STORER b On a.IMP_CODE=b.IMP_CODE and a.STORER_CODE=b.STORER_CODE Where a.IMP_CODE='" + IMP_CODE + "' and a.STORER_CODE='" + STORER_CODE + "' and ITM_CODE='" + ITEM_CODE + "' And PACK_KEY='1'"
            Dim dtItem As DataTable = gDB.getDataTable(SQLString, conn, transaction)

            SQLString = "Select a.CUS_NAME,a.CUS_CODE from WMS_CUSTOMER a Where a.IMP_CODE='" + IMP_CODE + "' and a.STORER_CODE='" + STORER_CODE + "' and CUS_CODE='" + CUS_CODE.Split(",")(0) + "' "
            Dim dtCust As DataTable = gDB.getDataTable(SQLString, conn, transaction)

            If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                SQLString = "Select a.*,IsNull(b.ITM_SHELF_LIFE,0)ITM_SHELF_LIFE from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 order by a.ILOC_EXPIRY_DATE "
            Else
                SQLString = "Select a.*,IsNull(b.ITM_SHELF_LIFE,0)ITM_SHELF_LIFE from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 order by a.ILOC_EXPIRY_DATE"
            End If
            dtStock = gDB.getDataTable(SQLString, conn, transaction)

            If dtStock IsNot Nothing And dtStock.Rows.Count > 0 Then
                If CUS_CODE <> "" Then
                    CUS_CODE = CUS_CODE.Split(",")(0)
                    SQLString = "Select ISNULL(MPL_FLAG,0)MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,ISNULL(MSL_FLAG,0)MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,ISNULL(MB_FLAG,0)MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,ISNULL(LOTS_CANNOT_BE_EARLIER,0)LOTS_CANNOT_BE_EARLIER,ISNULL(MAX_LOTS,0)MAX_LOTS from WMS_CUSTOMER_RULE where item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
                    dtRule = gDB.getDataTable(SQLString, conn, transaction)
                    If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
                        SQLString = "Select ISNULL(MPL_FLAG,0)MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,ISNULL(MSL_FLAG,0)MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,ISNULL(MB_FLAG,0)MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,ISNULL(LOTS_CANNOT_BE_EARLIER,0)LOTS_CANNOT_BE_EARLIER,ISNULL(MAX_LOTS,0)MAX_LOTS from WMS_CUSTOMER Where  IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
                        dtRule = gDB.getDataTable(SQLString, conn, transaction)
                    End If

                    If dtRule.Rows.Count > 0 AndAlso RULE_COUNT > 0 Then
                        Dim dtToday As DateTime = DO_DATE
                        Dim PrevLot As DateTime = Nothing
                        returnList.Clear()
                        IsFound = 0
                        Dim MAX_LOTS = Convert.ToInt32(dtRule.Rows(0)("MAX_LOTS").ToString())
                        For Each rowStk In dtStock.Rows


                            Dim dtExpiry As DateTime = Convert.ToDateTime(rowStk("ILOC_EXPIRY_DATE").ToString)
                            Dim strBatch As String = rowStk("ILOC_BATCH_NO").ToString
                            Dim dtLot As DateTime = Nothing
                            If strBatch IsNot Nothing AndAlso strBatch.Length >= 8 Then
                                dtLot = Convert.ToDateTime(strBatch.Substring(0, 4) + "-" + strBatch.Substring(4, 2) + "-" + strBatch.Substring(6, 2))
                            End If
                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "'  and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                            Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                                rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                            End If

                            If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                                If dtExpiry.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                    If dtLot >= Last_LOT_DT Then
                                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0  and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                                        Else
                                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                                        End If
                                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                            End If
                                                        Next
                                                        Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                                        If (drTemp.Length > 0) Then
                                                            dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                        Else
                                                            dtTmp = New DataTable()
                                                        End If
                                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                        Dim AvQty As Double = 0
                                                        If MAX_LOTS > 0 Then
                                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                Else
                                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                        If MAX_LOTS = 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If

                                                                        ElseIf MAX_LOTS = 2 Then

                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If

                                                                        ElseIf MAX_LOTS = 3 Then
                                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)

                                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)

                                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next
                                                                End If
                                                                If tmpList.Count > 0 Then
                                                                    Exit For
                                                                End If
                                                            Next
                                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    Dim balQty As Double = 0.00
                                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                        balQty = QTY

                                                                    Else
                                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                    End If
                                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                                    QTY = QTY - balQty
                                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                    returnList.Add(dictParam)
                                                                    IsFound = 1
                                                                    AllotedLotCount = AllotedLotCount + 1
                                                                Next
                                                            End If
                                                        Else
                                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                            Next
                                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY
                                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                                Else
                                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                                End If
                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                                'Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                                If dtLot >= Last_LOT_DT Then
                                                    If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 Order by a.ILOC_EXPIRY_DATE"
                                                    Else
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "'  and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 Order by a.ILOC_EXPIRY_DATE"
                                                    End If
                                                    Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                            dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                        End If
                                                    Next
                                                    Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                                    If (drTemp.Length > 0) Then
                                                        dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                    Else
                                                        dtTmp = New DataTable()
                                                    End If
                                                    Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                    Dim AvQty As Double = 0
                                                    If MAX_LOTS > 0 Then
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            Else
                                                                For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                    If MAX_LOTS = 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 2 Then

                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 3 Then
                                                                        If dtTmp.Rows.Count >= j + 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                            If tmpList.Count > 0 Then
                                                                Exit For
                                                            End If
                                                        Next
                                                        If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                            For Each item As Dictionary(Of String, Object) In tmpList
                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY

                                                                Else
                                                                    balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                End If
                                                                dictParam.Add("DOD_PK_QTY", balQty)

                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                            Next
                                                        End If
                                                    Else
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                        Next
                                                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY
                                                                dictParam.Add("DOD_PK_QTY", QTY)
                                                            Else
                                                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                            End If
                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                            'Exit For
                                                        End If
                                                    End If

                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                            If dtExpiry.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                                Else
                                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE ='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                                End If
                                                Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                    SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                    Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                    If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                        dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                    End If
                                                Next
                                                Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                                If (drTemp.Length > 0) Then
                                                    dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                Else
                                                    dtTmp = New DataTable()
                                                End If
                                                Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                Dim AvQty As Double = 0
                                                If MAX_LOTS > 0 Then
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                        If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)
                                                            Exit For
                                                        Else
                                                            For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                If MAX_LOTS = 1 Then
                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If

                                                                ElseIf MAX_LOTS = 2 Then

                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If

                                                                ElseIf MAX_LOTS = 3 Then
                                                                    If dtTmp.Rows.Count >= j + 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                        If tmpList.Count > 0 Then
                                                            Exit For
                                                        End If
                                                    Next
                                                    If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                        For Each item As Dictionary(Of String, Object) In tmpList
                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY

                                                            Else
                                                                balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                            End If
                                                            dictParam.Add("DOD_PK_QTY", balQty)

                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                        Next
                                                    End If
                                                Else
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                        AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                    Next
                                                    If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                        Dim dictParam As New Dictionary(Of String, Object)()
                                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        Dim balQty As Double = 0.00
                                                        If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                            balQty = QTY
                                                            dictParam.Add("DOD_PK_QTY", QTY)
                                                        Else
                                                            balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                            dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                        End If
                                                        QTY = QTY - balQty
                                                        Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                        returnList.Add(dictParam)
                                                        IsFound = 1
                                                        AllotedLotCount = AllotedLotCount + 1
                                                        'Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If LAST_LOT IsNot Nothing AndAlso LAST_LOT <> "" Then
                                    If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If dtExpiry.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                If dtLot >= Last_LOT_DT Then
                                                    If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order By a.ILOC_EXPIRY_DATE"
                                                    Else
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order By a.ILOC_EXPIRY_DATE"
                                                    End If
                                                    Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                            dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                        End If
                                                    Next
                                                    Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                                    If (drTemp.Length > 0) Then
                                                        dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                    Else
                                                        dtTmp = New DataTable()
                                                    End If
                                                    Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                    Dim AvQty As Double = 0
                                                    If MAX_LOTS > 0 Then
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            Else
                                                                For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                    If MAX_LOTS = 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 2 Then

                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 3 Then
                                                                        If dtTmp.Rows.Count >= j + 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                            If tmpList.Count > 0 Then
                                                                Exit For
                                                            End If
                                                        Next
                                                        If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                            For Each item As Dictionary(Of String, Object) In tmpList
                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY

                                                                Else
                                                                    balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                End If
                                                                dictParam.Add("DOD_PK_QTY", balQty)

                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                            Next
                                                        End If
                                                    Else
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                        Next
                                                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY
                                                                dictParam.Add("DOD_PK_QTY", QTY)
                                                            Else
                                                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                            End If
                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                            'Exit For
                                                        End If
                                                    End If
                                                    'PrevLot = dtLot
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If LAST_LOT.Length >= 8 Then
                                        Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                        If dtLot >= Last_LOT_DT Then
                                            If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                            Else
                                                SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                            End If
                                            Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                    dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                End If
                                            Next
                                            Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                            If (drTemp.Length > 0) Then
                                                dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                            Else
                                                dtTmp = New DataTable()
                                            End If
                                            Dim tmpList = New List(Of Dictionary(Of String, Object))
                                            Dim AvQty As Double = 0
                                            If MAX_LOTS > 0 Then
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                    If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    Else
                                                        For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                            If MAX_LOTS = 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If

                                                            ElseIf MAX_LOTS = 2 Then

                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If

                                                            ElseIf MAX_LOTS = 3 Then
                                                                If dtTmp.Rows.Count >= j + 1 Then
                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                    If tmpList.Count > 0 Then
                                                        Exit For
                                                    End If
                                                Next
                                                If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                    For Each item As Dictionary(Of String, Object) In tmpList
                                                        Dim dictParam As New Dictionary(Of String, Object)()
                                                        dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                        dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                        dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                        dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                        dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        Dim balQty As Double = 0.00
                                                        If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                            balQty = QTY

                                                        Else
                                                            balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                        End If
                                                        dictParam.Add("DOD_PK_QTY", balQty)

                                                        QTY = QTY - balQty
                                                        Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                        returnList.Add(dictParam)
                                                        IsFound = 1
                                                        AllotedLotCount = AllotedLotCount + 1
                                                    Next
                                                End If
                                            Else
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                    AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                                Next
                                                If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY
                                                        dictParam.Add("DOD_PK_QTY", QTY)
                                                    Else
                                                        balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                        dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                    End If
                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                    'Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtExpiry.Subtract(dtToday).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                        'Dim StkbalQty = Double.Parse(dtRSVD.Rows(0)("QTY"))
                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                        Else
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE) >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                        End If
                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                            End If
                                        Next
                                        Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                        If (drTemp.Length > 0) Then
                                            dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                        Else
                                            dtTmp = New DataTable()
                                        End If
                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                        Dim AvQty As Double = 0
                                        If MAX_LOTS > 0 Then
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                    tmpList.Add(tmpDict)
                                                    Exit For
                                                Else
                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                        If MAX_LOTS = 1 Then
                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 2 Then

                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)

                                                                tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 3 Then
                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                If tmpList.Count > 0 Then
                                                    Exit For
                                                End If
                                            Next
                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY

                                                    Else
                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                    End If
                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                Next
                                            End If
                                        Else
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                            Next
                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                Dim balQty As Double = 0.00
                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                    balQty = QTY
                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                Else
                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                End If
                                                QTY = QTY - balQty
                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                returnList.Add(dictParam)
                                                IsFound = 1
                                                AllotedLotCount = AllotedLotCount + 1
                                                'Exit For
                                            End If
                                        End If
                                        'PrevLot = dtLot
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If (dtToday - dtExpiry).Days <= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                        Else
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                        End If
                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                            End If
                                        Next
                                        Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                        If (drTemp.Length > 0) Then
                                            dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                        Else
                                            dtTmp = New DataTable()
                                        End If
                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                        Dim AvQty As Double = 0
                                        If MAX_LOTS > 0 Then
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                    tmpList.Add(tmpDict)
                                                    Exit For
                                                Else
                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                        If MAX_LOTS = 1 Then
                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 2 Then

                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)

                                                                tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 3 Then
                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                If tmpList.Count > 0 Then
                                                    Exit For
                                                End If
                                            Next
                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY

                                                    Else
                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                    End If
                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                Next
                                            End If
                                        Else
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                            Next
                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                Dim balQty As Double = 0.00
                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                    balQty = QTY
                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                Else
                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                End If
                                                QTY = QTY - balQty
                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                returnList.Add(dictParam)
                                                IsFound = 1
                                                AllotedLotCount = AllotedLotCount + 1
                                                'Exit For
                                            End If
                                        End If

                                    End If
                                End If
                            Else
                                'Dim StkbalQty = Double.Parse(dtRSVD.Rows(0)("QTY"))
                                If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                Else
                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1  Order by a.ILOC_EXPIRY_DATE"
                                End If
                                Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                For i As Integer = 0 To dtTmp.Rows.Count - 1
                                    SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                    Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                    If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                        dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                    End If
                                Next
                                Dim drTemp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE")
                                If (drTemp.Length > 0) Then
                                    dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                Else
                                    dtTmp = New DataTable()
                                End If
                                Dim tmpList = New List(Of Dictionary(Of String, Object))
                                Dim AvQty As Double = 0
                                If MAX_LOTS > 0 Then
                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                        If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                            Dim tmpDict = New Dictionary(Of String, Object)
                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                            tmpDict.Add("DOD_PK_QTY", 0)
                                            tmpList.Add(tmpDict)
                                            Exit For
                                        Else
                                            For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                If MAX_LOTS = 1 Then
                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    End If

                                                ElseIf MAX_LOTS = 2 Then

                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)

                                                        tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    End If

                                                ElseIf MAX_LOTS = 3 Then
                                                    If dtTmp.Rows.Count >= j + 1 Then
                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)

                                                            tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)

                                                            tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                        If tmpList.Count > 0 Then
                                            Exit For
                                        End If
                                    Next
                                    If QTY > 0 AndAlso tmpList.Count > 0 Then
                                        For Each item As Dictionary(Of String, Object) In tmpList
                                            Dim dictParam As New Dictionary(Of String, Object)()
                                            dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                            dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                            dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                            dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                            dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                            Dim balQty As Double = 0.00
                                            If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                balQty = QTY

                                            Else
                                                balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                            End If
                                            dictParam.Add("DOD_PK_QTY", balQty)

                                            QTY = QTY - balQty
                                            Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                            returnList.Add(dictParam)
                                            IsFound = 1
                                            AllotedLotCount = AllotedLotCount + 1
                                        Next
                                    End If
                                Else
                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                        AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                    Next
                                    If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                        Dim dictParam As New Dictionary(Of String, Object)()
                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                        Dim balQty As Double = 0.00
                                        If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                            balQty = QTY
                                            dictParam.Add("DOD_PK_QTY", QTY)
                                        Else
                                            balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                            dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                        End If
                                        QTY = QTY - balQty
                                        Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                        returnList.Add(dictParam)
                                        IsFound = 1
                                        AllotedLotCount = AllotedLotCount + 1
                                        'Exit For
                                    End If
                                End If

                            End If
                        Next
                        If IsFound = 0 Then
                            'lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                            updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                        End If
                    Else
                        IsFound = 0
                        Dim AvlQty As Double = 0.0000
                        For i As Integer = 0 To dtStock.Rows.Count - 1
                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtStock.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtStock.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtStock.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtStock.Rows(i)("ILOC_LOC")) & "'"
                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                dtStock.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                            End If
                            AvlQty = AvlQty + Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY"))
                        Next
                        For Each rowStk In dtStock.Rows
                            'SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "'  and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                            ''and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "'
                            'Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            'If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                            '    rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                            'End If
                            'If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                            '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 "
                            'Else
                            '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0  and a.ILOC_EXPIRY_DATE>=GETDATE()+1 "
                            'End If
                            'Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) >= 0 AndAlso AvlQty >= QTY Then 'AndAlso (Convert.ToDouble(dtTmp.Rows(0)("AvQty").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)) >= QTY Then
                                Dim dictParam As New Dictionary(Of String, Object)()
                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                dictParam.Add("ILOC_BAL_QTY", 0)
                                dictParam.Add("DOD_MIN_PROD_DATE", "0")
                                dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                                dictParam.Add("DOD_LAST_LOT", "")
                                dictParam.Add("DOD_MAX_LOT", "0")
                                Dim balQty As Double = 0.00

                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                    balQty = QTY
                                    dictParam.Add("DOD_PK_QTY", QTY)
                                Else
                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                End If
                                QTY = QTY - balQty
                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                returnList.Add(dictParam)
                                IsFound = 1


                                'Exit For

                            End If
                        Next
                        If IsFound = 0 Then
                            'lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                            updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                            'lstError.Add(ITEM_CODE + " Qty: " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " Not available in stock!")
                        End If
                    End If
                Else
                    IsFound = 0
                    Dim AvlQty As Double = 0.0000
                    For i As Integer = 0 To dtStock.Rows.Count - 1
                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtStock.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtStock.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtStock.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtStock.Rows(i)("ILOC_LOC")) & "'"
                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                            dtStock.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                        End If
                        AvlQty = AvlQty + Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY"))
                    Next
                    For Each rowStk In dtStock.Rows
                        'SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "' and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "' and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                        'Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                        'If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                        '    rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                        'End If

                        'If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                        '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 "
                        'Else
                        '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and a.ILOC_EXPIRY_DATE>=GETDATE()+1 "
                        'End If
                        'Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) >= 0 AndAlso AvlQty >= QTY Then ' (Convert.ToDouble(dtTmp.Rows(0)("AvQty").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)) >= QTY Then
                            Dim dictParam As New Dictionary(Of String, Object)()
                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                            dictParam.Add("ILOC_BAL_QTY", 0)
                            dictParam.Add("DOD_MIN_PROD_DATE", "0")
                            dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                            dictParam.Add("DOD_LAST_LOT", "")
                            dictParam.Add("DOD_MAX_LOT", "0")
                            Dim balQty As Double = 0.00

                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                balQty = QTY
                                dictParam.Add("DOD_PK_QTY", QTY)
                            Else
                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                            End If
                            QTY = QTY - balQty
                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                            returnList.Add(dictParam)
                            IsFound = 1


                            'Exit For

                        End If
                    Next
                    If IsFound = 0 Then
                        'lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                        updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                        'lstError.Add(ITEM_CODE + " Qty: " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " Not available in stock!")
                    End If
                End If

            Else
                'lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                'lstError.Add(ITEM_CODE + " Qty:  " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " not available in stock!")
                'Dim dictParam As New Dictionary(Of String, Object)()
                'dictParam.Add("ILOC_EXPIRY_DATE", "")
                'dictParam.Add("ITM_CODE", ITEM_CODE)
                'dictParam.Add("ILOC_BATCH_NO", "")
                'dictParam.Add("ILOC_LOC", "")
                'dictParam.Add("ILOC_WH", "")
                'dictParam.Add("ILOC_BAL_QTY", 0)
                'dictParam.Add("DOD_MIN_PROD_DATE", "0")
                'dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                'dictParam.Add("DOD_LAST_LOT", "")
                'dictParam.Add("DOD_MAX_LOT", "")
                'dictParam.Add("DOD_PK_QTY", "0")
                'SaveReserve(IMP_CODE, STORER_CODE, CUS_CODE, ITEM_CODE, "", QTY, ROUTE_ID, DO_DATE, DO_CODE)
                'returnList.Add(dictParam)
            End If
        Catch ex As Exception
            WriteExceptionLog(ex)
            Throw ex
        End Try
        Return returnList
    End Function

    Private Function SaveReserve(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal WH_LOC As String, ByVal WH_CODE As String, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal ITEM_CODE As String, ByVal LOT_NO As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime, ByVal DO_CODE As String) As Int16

        Dim sql_string As String = ""
        sql_string += "insert into WMS_WAVEPICK_RSVD (imp_code, storer_code, CUS_CODE,ITEM_CODE,LOT_NO,QTY,ROUTE_ID,DO_DATE,DO_CODE,WH_CODE,WH_LOC)"
        sql_string += "values ("
        sql_string += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
        sql_string += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(CUS_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ITEM_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(LOT_NO.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(QTY.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ROUTE_ID.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(DO_DATE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(DO_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(WH_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(WH_LOC.ToString.Trim) & "'"
        sql_string += ")"

        Dim RowCount = gDB.amendData(sql_string, conn, transaction)
        Return RowCount
    End Function

    Private Function updateCO_Details(ByVal ITEM_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal DO_CODE As String, ByVal ROUTE_ID As String, ByRef conn As SqlConnection, ByRef transaction As SqlTransaction) As String
        Dim UpdateSql As String
        Dim SQLString As String
        Dim cocodes As String = ""

        SQLString = "Select DOD_CO_CODE from WMS_DELV_ORDER_D where DO_CODE='" & DO_CODE & "'"
        Dim dtcoCodes = gDB.getDataTable(SQLString, conn, transaction)
        If (dtcoCodes.Rows.Count = 0) Then
            UpdateSql = "Update cod set cod.ALLOCATED='N' FROM WMS_CUST_ORDER_D cod Inner JOIN WMS_CUST_ORDER co ON co.CO_CODE=cod.CO_CODE where cod.COD_ITM_CODE = '" + ITEM_CODE + "' and cod.STORER_CODE = '" + STORER_CODE + "' and co.CUS_CODE='" + CUS_CODE + "' and co.CO_STATUS='NEW' and co.ROUTE_ID='" + ROUTE_ID + "'"
            gDB.amendData(UpdateSql, conn, transaction)
        End If
        For Each row As DataRow In dtcoCodes.Rows
            Dim co_Codes = row("DOD_CO_CODE").ToString().Split(",")
            For Each cocd As String In co_Codes
                If cocodes = "" Then
                    cocodes = "" + cocd + ""
                Else
                    cocodes = cocodes + "," + cocd + ""
                End If
            Next
        Next

        UpdateSql = "Update cod set cod.ALLOCATED='N' FROM WMS_CUST_ORDER_D cod Inner JOIN WMS_CUST_ORDER co ON co.CO_CODE=cod.CO_CODE where cod.CO_CODE in (select items from  dbo.split('" & cocodes & "',',')) and cod.COD_ITM_CODE = '" + ITEM_CODE + "' and cod.STORER_CODE = '" + STORER_CODE + "' and co.CUS_CODE='" + CUS_CODE + "' and co.ROUTE_ID='" + ROUTE_ID + "'"
        gDB.amendData(UpdateSql, conn, transaction)
        Return True
    End Function

    'Import ITEM MASTER
    Private Sub ImportITEMMASTERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select IO_ID,ITEM_ID,ITEM_NUMBER,ITEM_STATUS_CODE,ITEM_DESCRIPTION,ITEM_DESCRIPTION_ZHS,ITEM_TYPE,SHELF_LIFE_DAYS,PRIMARY_UOM_CODE,SEC_UOM_CODE,SEC_UOM_CONV,LOT_EXISTS,INSPECTION_REQUIRED_FLAG,ITEM_PRODUCT_CAT_DESC,ITEM_BRAND_DESC,MIN_SHELF_LIFE_DAYS,RELATED_ITEM_ID,RELATED_ITEM_NUMBER,IsNUll(SUBINVENTORY_CODE,'FG01')SUBINVENTORY_CODE from EBS_WMS_ITEM_MASTER"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                Dim ITEMInsert As Int16 = 0
                For Each rowD As DataRow In codt.Rows
                    SQLString = "Select * from WMS_ITEM " &
                        "WHERE STORER_CODE = '" & (rowD("IO_ID").ToString.Trim) & "'and IMP_CODE = 'WMS' and ITM_CODE='" & (rowD("ITEM_ID").ToString.Trim) & "' and PACK_KEY='1'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        SQLString = "Insert into WMS_ITEM(IMP_CODE,STORER_CODE,ITM_CODE,PACK_KEY,ITM_SKU_NO,ITM_STATUS,ITM_NAME,ITM_NAME_CH,ITM_DESC,ITM_DESC_CH,ITM_TYPE,ITM_SHELF_LIFE,ITM_UOM,ITM_CAT,ITM_BRAND,ITM_STACKABLE_YN,ITM_INSP_YN,ITM_UOM2,ITM_QTY2,ITM_MAX_STOCK,ITM_MFG,ITM_DRAWING_NO,ITM_PREF_WH,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE, @STORER_CODE,@ITM_CODE,@PACK_KEY,@ITM_SKU_NO,@ITM_STATUS,@ITM_NAME,@ITM_NAME_CH,@ITM_DESC,@ITM_DESC_CH,@ITM_TYPE,@ITM_SHELF_LIFE,@ITM_UOM,@ITM_CAT,@ITM_BRAND,@ITM_STACKABLE_YN,@ITM_INSP_YN,@ITM_UOM2,@ITM_QTY2,@ITM_MAX_STOCK,@ITM_MFG,@ITM_DRAWING_NO,@ITM_PREF_WH,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", rowD("IO_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_CODE", rowD("ITEM_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@PACK_KEY", "1")
                        cmd.Parameters.AddWithValue("@ITM_SKU_NO", rowD("ITEM_NUMBER").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_STATUS", rowD("ITEM_STATUS_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_NAME", rowD("ITEM_DESCRIPTION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_NAME_CH", rowD("ITEM_DESCRIPTION_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DESC", rowD("ITEM_DESCRIPTION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DESC_CH", rowD("ITEM_DESCRIPTION_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_TYPE", rowD("ITEM_TYPE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_SHELF_LIFE", rowD("SHELF_LIFE_DAYS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_UOM", rowD("PRIMARY_UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_CAT", rowD("ITEM_PRODUCT_CAT_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_BRAND", rowD("ITEM_BRAND_DESC").ToString.Trim)
                        If (rowD("LOT_EXISTS").ToString.Trim) = "1" Then
                            cmd.Parameters.AddWithValue("@ITM_STACKABLE_YN", "Y")
                        ElseIf (rowD("LOT_EXISTS").ToString.Trim) = "0" Then
                            cmd.Parameters.AddWithValue("@ITM_STACKABLE_YN", "N")
                        Else
                            cmd.Parameters.AddWithValue("@ITM_STACKABLE_YN", "N")
                        End If
                        'cmd.Parameters.AddWithValue("@ITM_STACKABLE_YN", rowD("LOT_EXISTS").ToString.Trim)
                        If (rowD("INSPECTION_REQUIRED_FLAG").ToString.Trim) = "1" Then
                            cmd.Parameters.AddWithValue("@ITM_INSP_YN", "Y")
                        ElseIf (rowD("INSPECTION_REQUIRED_FLAG").ToString.Trim) = "0" Then
                            cmd.Parameters.AddWithValue("@ITM_INSP_YN", "N")
                        Else
                            cmd.Parameters.AddWithValue("@ITM_INSP_YN", "N")
                        End If
                        'cmd.Parameters.AddWithValue("@ITM_INSP_YN", rowD("INSPECTION_REQUIRED_FLAG").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_UOM2", rowD("SEC_UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_QTY2", rowD("SEC_UOM_CONV").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_MAX_STOCK", rowD("MIN_SHELF_LIFE_DAYS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_MFG", rowD("RELATED_ITEM_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DRAWING_NO", rowD("RELATED_ITEM_NUMBER").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_PREF_WH", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1

                        ITEMInsert = 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_ITEM Set ")
                        sbCmdText.Append("ITM_STATUS = @ITM_STATUS,")
                        sbCmdText.Append("ITM_NAME = @ITM_NAME,")
                        sbCmdText.Append("ITM_NAME_CH = @ITM_NAME_CH,")
                        sbCmdText.Append("ITM_DESC = @ITM_DESC,")
                        sbCmdText.Append("ITM_DESC_CH = @ITM_DESC_CH,")
                        sbCmdText.Append("ITM_SKU_NO = @ITM_SKU_NO,")
                        sbCmdText.Append("ITM_TYPE = @ITM_TYPE,")
                        sbCmdText.Append("ITM_SHELF_LIFE = @ITM_SHELF_LIFE,")
                        sbCmdText.Append("ITM_UOM = @ITM_UOM,")
                        sbCmdText.Append("ITM_CAT = @ITM_CAT,")
                        sbCmdText.Append("ITM_BRAND = @ITM_BRAND,")
                        sbCmdText.Append("ITM_UOM2 = @ITM_UOM2,")
                        sbCmdText.Append("ITM_QTY2 = @ITM_QTY2,")
                        sbCmdText.Append("ITM_MAX_STOCK = @ITM_MAX_STOCK,")
                        sbCmdText.Append("ITM_MFG = @ITM_MFG,")
                        sbCmdText.Append("ITM_DRAWING_NO = @ITM_DRAWING_NO,")
                        sbCmdText.Append("ITM_PREF_WH = @ITM_PREF_WH,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where STORER_CODE=@STORER_CODE and IMP_CODE=@IMP_CODE and ITM_CODE=@ITM_CODE and PACK_KEY=@PACK_KEY")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", rowD("IO_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_CODE", rowD("ITEM_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@PACK_KEY", "1")
                        cmd.Parameters.AddWithValue("@ITM_SKU_NO", rowD("ITEM_NUMBER").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_STATUS", rowD("ITEM_STATUS_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_NAME", rowD("ITEM_DESCRIPTION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_NAME_CH", rowD("ITEM_DESCRIPTION_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DESC", rowD("ITEM_DESCRIPTION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DESC_CH", rowD("ITEM_DESCRIPTION_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_TYPE", rowD("ITEM_TYPE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_SHELF_LIFE", rowD("SHELF_LIFE_DAYS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_UOM", rowD("PRIMARY_UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_CAT", rowD("ITEM_PRODUCT_CAT_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_BRAND", rowD("ITEM_BRAND_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_UOM2", rowD("SEC_UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_QTY2", rowD("SEC_UOM_CONV").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_MAX_STOCK", rowD("MIN_SHELF_LIFE_DAYS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_MFG", rowD("RELATED_ITEM_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_DRAWING_NO", rowD("RELATED_ITEM_NUMBER").ToString.Trim)
                        cmd.Parameters.AddWithValue("@ITM_PREF_WH", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCount = updateCount + 1

                        'lblMSG.Text = "ITEM MASTER data has been already updated"
                    End If

                Next

                'For INTERFACE LOG
                If ITEMInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "ITEM_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_ITEM_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_ITEM"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    ITFSTATUS = "SUCCESS"

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='ITEM MASTER'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", "ITEM MASTER", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted Item Master Record : " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated Item Master Record : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>ITEM MASTER data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "ITEM MASTER data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

    End Sub

    'Import CUSTOMER MASTER
    Private Sub ImportCUSTOMERMASTERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select DIVISION_CODE,ACCOUNT_NUMBER,IsNUll(CUSTOMER_NAME,ACCOUNT_NUMBER)CUSTOMER_NAME,IsNUll(CUSTOMER_NAME_ZHS,ACCOUNT_NUMBER)CUSTOMER_NAME_ZHS,ADDRESS1,ADDRESS2,ADDRESS3,ADDRESS4,SITE_USE_CODE,(case when isnull(MAX_LOTS,'')='' then 0 else MAX_LOTS end) as MAX_LOTS,(case when isnull(LOT_NOTEARLY,'')='Y' then 1 else 0 end) as LOT_NOTEARLY from EBS_WMS_CUSTOMER_MASTER "
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                Dim CUSInsert As Int16 = 0

                For Each rowD As DataRow In codt.Rows
                    Dim Storer As String
                    SQLString = "Select STORER_CODE from WMS_STORER " &
                    "WHERE DIVISION_CODE='" & (rowD("DIVISION_CODE").ToString.Trim) & "'"
                    Dim dtStorer As DataTable = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtStorer Is Nothing Or dtStorer.Rows.Count <= 0 Then
                        Storer = ""
                    Else
                        Storer = dtStorer.Rows(0)("STORER_CODE").ToString
                    End If

                    SQLString = "Select * from WMS_CUSTOMER " &
                        "WHERE STORER_CODE = '" + Storer + "' and IMP_CODE = 'WMS' and CUS_CODE = '" & (rowD("ACCOUNT_NUMBER").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)

                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_CUSTOMER(IMP_CODE,STORER_CODE,CUS_CODE,CUS_STATUS,CUS_NAME,CUS_NAME_CH,CUS_ADDR1_DEL,CUS_ADDR2_DEL,CUS_ADDR3_DEL,CUS_AREA_DEL,CUS_ADDR1_BILL,CUS_ADDR2_BILL,CUS_ADDR3_BILL,CUS_AREA_BILL,SHIP_TO,BILL_TO,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD,MAX_LOTS,LOTS_CANNOT_BE_EARLIER) VALUES (@IMP_CODE, @STORER_CODE,@CUS_CODE,@CUS_STATUS,@CUS_NAME,@CUS_NAME_CH,@CUS_ADDR1_DEL,@CUS_ADDR2_DEL,@CUS_ADDR3_DEL,@CUS_AREA_DEL,@CUS_ADDR1_BILL,@CUS_ADDR2_BILL,@CUS_ADDR3_BILL,@CUS_AREA_BILL,@SHIP_TO,@BILL_TO,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD,@MAX_LOTS,@LOTS_CANNOT_BE_EARLIER)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", Storer)
                        cmd.Parameters.AddWithValue("@CUS_CODE", rowD("ACCOUNT_NUMBER").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_STATUS", "ACTIVE")
                        cmd.Parameters.AddWithValue("@CUS_NAME", rowD("CUSTOMER_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_NAME_CH", rowD("CUSTOMER_NAME_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR1_DEL", rowD("ADDRESS1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR2_DEL", rowD("ADDRESS2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR3_DEL", rowD("ADDRESS3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_AREA_DEL", rowD("ADDRESS4").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR1_BILL", rowD("ADDRESS1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR2_BILL", rowD("ADDRESS2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR3_BILL", rowD("ADDRESS3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_AREA_BILL", rowD("ADDRESS4").ToString.Trim)

                        If (rowD("SITE_USE_CODE").ToString.Trim = "SHIP_TO") Then
                            cmd.Parameters.AddWithValue("@SHIP_TO", "SHIP_TO")
                        Else
                            cmd.Parameters.AddWithValue("@SHIP_TO", "")
                        End If

                        If (rowD("SITE_USE_CODE").ToString.Trim = "BILL_TO") Then
                            cmd.Parameters.AddWithValue("@BILL_TO", "BILL_TO")
                        Else
                            cmd.Parameters.AddWithValue("@BILL_TO", "")
                        End If

                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@MAX_LOTS", rowD("MAX_LOTS"))
                        cmd.Parameters.AddWithValue("@LOTS_CANNOT_BE_EARLIER", rowD("LOT_NOTEARLY"))
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1
                        CUSInsert = 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_CUSTOMER Set ")
                        'sbCmdText.Append("CUS_STATUS = @CUS_STATUS,")
                        sbCmdText.Append("CUS_NAME = @CUS_NAME,")
                        sbCmdText.Append("CUS_NAME_CH = @CUS_NAME_CH,")
                        sbCmdText.Append("CUS_ADDR1_DEL = @CUS_ADDR1_DEL,")
                        sbCmdText.Append("CUS_ADDR2_DEL = @CUS_ADDR2_DEL,")
                        sbCmdText.Append("CUS_ADDR3_DEL = @CUS_ADDR3_DEL,")
                        sbCmdText.Append("CUS_AREA_DEL = @CUS_AREA_DEL,")
                        sbCmdText.Append("CUS_ADDR1_BILL = @CUS_ADDR1_BILL,")
                        sbCmdText.Append("CUS_ADDR2_BILL = @CUS_ADDR2_BILL,")
                        sbCmdText.Append("CUS_ADDR3_BILL = @CUS_ADDR3_BILL,")
                        sbCmdText.Append("CUS_AREA_BILL = @CUS_AREA_BILL,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("MAX_LOTS = @MAX_LOTS,")
                        sbCmdText.Append("LOTS_CANNOT_BE_EARLIER = @LOTS_CANNOT_BE_EARLIER,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where STORER_CODE=@STORER_CODE and IMP_CODE=@IMP_CODE and CUS_CODE=@CUS_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", Storer)
                        cmd.Parameters.AddWithValue("@CUS_CODE", rowD("ACCOUNT_NUMBER").ToString.Trim)
                        'cmd.Parameters.AddWithValue("@CUS_STATUS", "ACTIVE")
                        cmd.Parameters.AddWithValue("@CUS_NAME", rowD("CUSTOMER_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_NAME_CH", rowD("CUSTOMER_NAME_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR1_DEL", rowD("ADDRESS1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR2_DEL", rowD("ADDRESS2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR3_DEL", rowD("ADDRESS3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_AREA_DEL", rowD("ADDRESS4").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR1_BILL", rowD("ADDRESS1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR2_BILL", rowD("ADDRESS2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_ADDR3_BILL", rowD("ADDRESS3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@CUS_AREA_BILL", rowD("ADDRESS4").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@MAX_LOTS", rowD("MAX_LOTS"))
                        cmd.Parameters.AddWithValue("@LOTS_CANNOT_BE_EARLIER", rowD("LOT_NOTEARLY"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCount = updateCount + 1

                        'lblMSG.Text = "CUSTOMER MASTER data has been already updated"
                    End If

                Next

                'For INTERFACE LOG
                If CUSInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "CUSTOMER_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_CUSTOMER_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_CUSTOMER"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    ITFSTATUS = "SUCCESS"

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='CUSTOMER MASTER'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", "CUSTOMER MASTER", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted Customer Master Record  : " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated Customer Master Record  : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>CUSTOMER MASTER data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "CUSTOMER MASTER data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import UOM MASTER
    Private Sub ImportUOMMASTERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select UOM_CODE,UOM_DESC from EBS_WMS_UOM_MASTER"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                Dim UOMInsert As Int16 = 0

                For Each rowD As DataRow In codt.Rows
                    SQLString = "Select * from WMS_UOM " &
                        "WHERE UOM_CODE = '" & (rowD("UOM_CODE").ToString.Trim) & "' and IMP_CODE = 'WMS'"

                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_UOM(IMP_CODE,UOM_CODE,UOM_DESC,SYS_LUB,SYS_CB,SYS_CD,SYS_LUD) VALUES (@IMP_CODE, @UOM_CODE,@UOM_DESC,@SYS_LUB,@SYS_CB,@SYS_CD,@SYS_LUD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@UOM_CODE", rowD("UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@UOM_DESC", rowD("UOM_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1
                        UOMInsert = 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_UOM Set ")
                        sbCmdText.Append("UOM_DESC = @UOM_DESC,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where UOM_CODE=@UOM_CODE and IMP_CODE=@IMP_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@UOM_CODE", rowD("UOM_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@UOM_DESC", rowD("UOM_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCount = updateCount + 1

                        'lblMSG.Text = "UOM MASTER data has been already updated"
                    End If
                Next

                'For INTERFACE LOG
                If UOMInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "UOM_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_UOM_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_UOM"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    ITFSTATUS = "SUCCESS"

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='UOM MASTER'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", "UOM MASTER", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted UOM Master Record : " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated UOM Master Record : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>UOM MASTER data has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "UOM MASTER data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import WAREHOUSE MASTER
    Private Sub ImportWAREHOUSEMASTERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            Dim TodayDateTime As DateTime
            TodayDateTime = System.DateTime.Now

            SQLString = "Select SUBINVENTORY_CODE,SUB_DESC,SUBINVENTORY_TYPE from EBS_WMS_WAREHOUSE_MASTER"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                Dim WAREHOUSEInsert As Int16 = 0

                For Each rowD As DataRow In codt.Rows
                    SQLString = "Select * from WMS_WAREHOUSE " &
                        "WHERE IMP_CODE='WMS' and WH_CODE = '" & (rowD("SUBINVENTORY_CODE").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)

                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_WAREHOUSE(IMP_CODE,WH_CODE,WH_NAME,WH_MAIN_WH,WH_TYPE,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE,@WH_CODE,@WH_NAME,@WH_MAIN_WH,@WH_TYPE,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        'cmd.Parameters.AddWithValue("@WH_OWNER", rowD("IO_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_NAME", rowD("SUB_DESC").ToString.Trim)
                        'cmd.Parameters.AddWithValue("@STATUS_ID", rowD("STATUS_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_MAIN_WH", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_TYPE", rowD("SUBINVENTORY_TYPE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1
                        WAREHOUSEInsert = 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_WAREHOUSE Set ")
                        sbCmdText.Append("WH_NAME = @WH_NAME,")
                        sbCmdText.Append("WH_MAIN_WH = @WH_MAIN_WH,")
                        sbCmdText.Append("WH_TYPE = @WH_TYPE,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and WH_CODE=@WH_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_NAME", rowD("SUB_DESC").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_MAIN_WH", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@WH_TYPE", rowD("SUBINVENTORY_TYPE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCount = updateCount + 1

                        'lblMSG.Text = "Warehouse MASTER data has been already updated"
                    End If

                Next

                'For INTERFACE LOG
                If WAREHOUSEInsert = 1 Then
                    Dim IMP_FILE_NAME As String
                    IMP_FILE_NAME = "WAREHOUSE_IMP_WMS_EBS_FILE"

                    'FOR INT_BATCH_NO'
                    Dim INT_BATCH_NO As String
                    INT_BATCH_NO = "WMS_WAREHOUSE_BATCH_NO"
                    'INT_BATCH_NO = TodayDateTime
                    'If INT_BATCH_NO <> "" Then
                    '    Dim d As DateTime = DateTime.ParseExact(INT_BATCH_NO, "dd-MMM-yyyy", CultureInfo.InvariantCulture)
                    '    Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    '    INT_BATCH_NO = reformatted
                    'End If

                    'for get current datetime'
                    Dim dtCurDateTime, dtTotalRecords As DataTable
                    SQLString = "select GETDATE() as CURRENTDATETIME"
                    dtCurDateTime = gDB.getDataTable(SQLString, gConn, transaction)

                    'for get total records of REPLENISH(PO IMPORT) table'
                    SQLString = "select count(*) as TotalRecords from WMS_WAREHOUSE"
                    dtTotalRecords = gDB.getDataTable(SQLString, gConn, transaction)

                    Dim ITFSTATUS As String
                    ITFSTATUS = "SUCCESS"

                    Dim dtINTFLOG As New DataTable
                    SQLString = "select TOP(1) * FROM WMS_INTF_LOG WHERE ITF_IMP_TYPE='WAREHOUSE MASTER'"
                    dtINTFLOG = gDB.getDataTable(SQLString, gConn, transaction)
                    If dtINTFLOG Is Nothing Or dtINTFLOG.Rows.Count <= 0 Then
                        'call InterfaceLog method'
                        InterfaceLog(INT_BATCH_NO, IMP_FILE_NAME, "I", "WAREHOUSE MASTER", TodayDateTime, dtCurDateTime.Rows(0)("CURRENTDATETIME").ToString.Trim, ITFSTATUS, "", "", dtTotalRecords.Rows(0)("TotalRecords").ToString.Trim(), insertCount, 0)
                    End If
                End If

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted Warehouse Master Record : " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated Warehouse Master Record : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>Warehouse Master has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "Warehouse MASTER data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import WAREHOUSE LOCATION
    Private Sub ImportWAREHOUSELOCATIONData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCountFL As Int16 = 0
        Dim updateCountFL As Int16 = 0
        Dim insertCountAR As Int16 = 0
        Dim updateCountAR As Int16 = 0
        Dim insertCountRK As Int16 = 0
        Dim updateCountRK As Int16 = 0
        Dim insertCountBN As Int16 = 0
        Dim updateCountBN As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            SQLString = "Select SUBINVENTORY_CODE,IsNUll( DESCRIPTION,SUBINVENTORY_CODE)DESCRIPTION,SEGMENT1,SEGMENT2,SEGMENT3,SEGMENT4 from EBS_WMS_WAREHOUSE_LOCATION"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then

                For Each rowD As DataRow In codt.Rows
                    'For FLOOR
                    SQLString = "Select * from WMS_WH_FL " &
                        "WHERE IMP_CODE='WMS' and WH_CODE = '" & (rowD("SUBINVENTORY_CODE").ToString.Trim) & "' and FL_NUM='" & (rowD("SEGMENT1").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_WH_FL(IMP_CODE,WH_CODE,FL_NUM,FL_NAME,FL_NAME_CH,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE,@WH_CODE,@FL_NUM,@FL_NAME,@FL_NAME_CH,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NAME", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NAME_CH", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCountFL = insertCountFL + 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_WH_FL Set ")
                        sbCmdText.Append("FL_NAME = @FL_NAME,")
                        sbCmdText.Append("FL_NAME_CH = @FL_NAME_CH,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and WH_CODE=@WH_CODE and FL_NUM=@FL_NUM")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NAME", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NAME_CH", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCountFL = updateCountFL + 1
                    End If

                    'For AREA
                    SQLString = "Select * from WMS_WH_AREA " &
                        "WHERE IMP_CODE='WMS' and WH_CODE = '" & (rowD("SUBINVENTORY_CODE").ToString.Trim) & "' and FL_NUM ='" & (rowD("SEGMENT1").ToString.Trim) & "' and AR_CODE ='" & (rowD("SEGMENT2").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_WH_AREA(IMP_CODE,WH_CODE,FL_NUM,AR_CODE,AR_NAME,AR_NAME_CH,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE,@WH_CODE,@FL_NUM,@AR_CODE,@AR_NAME,@AR_NAME_CH,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_NAME", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_NAME_CH", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCountAR = insertCountAR + 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_WH_AREA Set ")
                        sbCmdText.Append("AR_NAME = @AR_NAME,")
                        sbCmdText.Append("AR_NAME_CH = @AR_NAME_CH,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and WH_CODE=@WH_CODE and FL_NUM = @FL_NUM and AR_CODE=@AR_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_NAME", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_NAME_CH", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCountAR = updateCountAR + 1
                    End If

                    'For RACK
                    SQLString = "Select * from WMS_WH_RACK " &
                        "WHERE IMP_CODE='WMS' and WH_CODE = '" & (rowD("SUBINVENTORY_CODE").ToString.Trim) & "' and FL_NUM='" & (rowD("SEGMENT1").ToString.Trim) & "' and AR_CODE='" & (rowD("SEGMENT2").ToString.Trim) & "' and RK_CODE='" & (rowD("SEGMENT3").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_WH_RACK(IMP_CODE,WH_CODE,FL_NUM,AR_CODE,RK_CODE,RK_NAME,RK_NAME_CH,RK_X,RK_Y,RK_XBINS,RK_YBINS,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE,@WH_CODE,@FL_NUM,@AR_CODE,@RK_CODE,@RK_NAME,@RK_NAME_CH,@RK_X,@RK_Y,@RK_XBINS,@RK_YBINS,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_CODE", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_NAME", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_NAME_CH", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_X", 1)
                        cmd.Parameters.AddWithValue("@RK_Y", 1)
                        cmd.Parameters.AddWithValue("@RK_XBINS", 1)
                        cmd.Parameters.AddWithValue("@RK_YBINS", 1)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCountRK = insertCountRK + 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_WH_RACK Set ")
                        sbCmdText.Append("RK_NAME = @RK_NAME,")
                        sbCmdText.Append("RK_NAME_CH = @RK_NAME_CH,")
                        sbCmdText.Append("RK_X = @RK_X,")
                        sbCmdText.Append("RK_Y = @RK_Y,")
                        sbCmdText.Append("RK_XBINS = @RK_XBINS,")
                        sbCmdText.Append("RK_YBINS = @RK_YBINS,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and WH_CODE=@WH_CODE and FL_NUM = @FL_NUM and AR_CODE = @AR_CODE and RK_CODE = @RK_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_CODE", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_NAME", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_NAME_CH", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_X", 1)
                        cmd.Parameters.AddWithValue("@RK_Y", 1)
                        cmd.Parameters.AddWithValue("@RK_XBINS", 1)
                        cmd.Parameters.AddWithValue("@RK_YBINS", 1)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCountRK = updateCountRK + 1
                    End If

                    'For BIN
                    SQLString = "Select * from WMS_WH_BIN " &
                        "WHERE IMP_CODE='WMS' and WH_CODE = '" & (rowD("SUBINVENTORY_CODE").ToString.Trim) & "' and FL_NUM='" & (rowD("SEGMENT1").ToString.Trim) & "' and AR_CODE='" & (rowD("SEGMENT2").ToString.Trim) & "' and RK_CODE='" & (rowD("SEGMENT3").ToString.Trim) & "' and BN_CODE='" & (rowD("SEGMENT4").ToString.Trim) & "'"
                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    Dim lOC = rowD("SEGMENT1").ToString.Trim + rowD("SEGMENT2").ToString.Trim + rowD("SEGMENT3").ToString.Trim + rowD("SEGMENT4").ToString.Trim
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_WH_BIN(IMP_CODE,WH_CODE,FL_NUM,AR_CODE,RK_CODE,BN_CODE,BN_X,BN_Y,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE,@WH_CODE,@FL_NUM,@AR_CODE,@RK_CODE,@BN_CODE,@BN_X,@BN_Y,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        'LOC_KEY,@LOC_KEY
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_CODE", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@BN_CODE", rowD("SEGMENT4").ToString.Trim)
                        cmd.Parameters.AddWithValue("@BN_X", 1)
                        cmd.Parameters.AddWithValue("@BN_Y", 1)
                        'cmd.Parameters.AddWithValue("@LOC_KEY", lOC)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCountBN = insertCountBN + 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_WH_BIN Set ")
                        sbCmdText.Append("BN_CODE = @BN_CODE,")
                        sbCmdText.Append("BN_X = @BN_X,")
                        sbCmdText.Append("BN_Y = @BN_Y,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and WH_CODE=@WH_CODE and FL_NUM = @FL_NUM and AR_CODE = @AR_CODE and RK_CODE = @RK_CODE and BN_CODE=@BN_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@WH_CODE", rowD("SUBINVENTORY_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@FL_NUM", rowD("SEGMENT1").ToString.Trim)
                        cmd.Parameters.AddWithValue("@AR_CODE", rowD("SEGMENT2").ToString.Trim)
                        cmd.Parameters.AddWithValue("@RK_CODE", rowD("SEGMENT3").ToString.Trim)
                        cmd.Parameters.AddWithValue("@BN_CODE", rowD("SEGMENT4").ToString.Trim)
                        cmd.Parameters.AddWithValue("@BN_X", 1)
                        cmd.Parameters.AddWithValue("@BN_Y", 1)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCountBN = updateCountBN + 1
                    End If
                Next

            End If

            transaction.Commit()

            Dim strAlert As String = "<p>Row Inserted of Floor : " & insertCountFL.ToString.Trim & "</p>"
            strAlert &= "<p>Row Updated of Floor : " & updateCountFL.ToString.Trim & "</p>"
            strAlert &= "<p>Row Inserted of Area : " & insertCountAR.ToString.Trim & "</p>"
            strAlert &= "<p>Row Updated of Area : " & updateCountAR.ToString.Trim & "</p>"
            strAlert &= "<p>Row Inserted of Rack : " & insertCountRK.ToString.Trim & "</p>"
            strAlert &= "<p>Row Updated of Rack : " & updateCountRK.ToString.Trim & "</p>"
            strAlert &= "<p>Row Inserted of Bin : " & insertCountBN.ToString.Trim & "</p>"
            strAlert &= "<p>Row Updated of Bin : " & updateCountBN.ToString.Trim & "</p>"
            'strAlert &= "<p>Warehouse Location data has been successfully imported!!</p>"
            lblMSG.Text = strAlert

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

    End Sub

    'Import STORER
    Private Sub ImportSTORERData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim codtl As DataTable
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            SQLString = "Select IO_ID,IO_NAME,COMPANY_NAME,COMPANY_NAME_ZHS,REGION,DIVISION_CODE,DIVISION_NAME from EBS_WMS_COMPANY_MASTER"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)
            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then

                For Each rowD As DataRow In codt.Rows
                    SQLString = "Select * from WMS_STORER " &
                        "WHERE IMP_CODE='WMS' and STORER_CODE = '" & (rowD("IO_ID").ToString.Trim) & "'"

                    codtl = gDB.getDataTable(SQLString, gConn, transaction)
                    If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                        'INSERT
                        'Parameterized Parameter
                        SQLString = "Insert into WMS_STORER(IMP_CODE,STORER_CODE,STO_STATUS,STO_SHORTNAME,STO_NAME,STO_NAME_CH,STO_REGION,STO_CAT,DIVISION_CODE,SYS_LUB,SYS_CB,SYS_LUD,SYS_CD) VALUES (@IMP_CODE, @STORER_CODE,@STO_STATUS,@STO_SHORTNAME,@STO_NAME,@STO_NAME_CH,@STO_REGION,@STO_CAT,@DIVISION_CODE,@SYS_LUB,@SYS_CB,@SYS_LUD,@SYS_CD)"
                        cmd = New SqlCommand(SQLString, gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", rowD("IO_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_STATUS", "ACTIVE")
                        cmd.Parameters.AddWithValue("@STO_SHORTNAME", rowD("IO_NAME").ToString.Trim + "(" + rowD("DIVISION_NAME").ToString.Trim + ")")
                        cmd.Parameters.AddWithValue("@STO_NAME", rowD("COMPANY_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_NAME_CH", rowD("COMPANY_NAME_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_REGION", rowD("REGION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@DIVISION_CODE", rowD("DIVISION_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_CAT", rowD("DIVISION_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                        cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        insertCount = insertCount + 1
                    Else
                        'UPDATE
                        sbCmdText = New StringBuilder()
                        sbCmdText.Append("Update WMS_STORER Set ")
                        sbCmdText.Append("STO_STATUS = @STO_STATUS,")
                        sbCmdText.Append("STO_SHORTNAME = @STO_SHORTNAME,")
                        sbCmdText.Append("STO_NAME = @STO_NAME,")
                        sbCmdText.Append("STO_NAME_CH = @STO_NAME_CH,")
                        sbCmdText.Append("DIVISION_CODE = @DIVISION_CODE,")
                        sbCmdText.Append("STO_CAT = @STO_CAT,")
                        sbCmdText.Append("STO_REGION = @STO_REGION,")
                        sbCmdText.Append("SYS_LUB = @SYS_LUB,")
                        sbCmdText.Append("SYS_LUD = @SYS_LUD")
                        sbCmdText.Append(" Where IMP_CODE=@IMP_CODE and STORER_CODE=@STORER_CODE")
                        cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                        cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                        cmd.Parameters.AddWithValue("@STORER_CODE", rowD("IO_ID").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_STATUS", "ACTIVE")
                        cmd.Parameters.AddWithValue("@STO_SHORTNAME", rowD("IO_NAME").ToString.Trim + "(" + rowD("DIVISION_NAME").ToString.Trim + ")")
                        cmd.Parameters.AddWithValue("@STO_NAME", rowD("COMPANY_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_NAME_CH", rowD("COMPANY_NAME_ZHS").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_REGION", rowD("REGION").ToString.Trim)
                        cmd.Parameters.AddWithValue("@DIVISION_CODE", rowD("DIVISION_CODE").ToString.Trim)
                        cmd.Parameters.AddWithValue("@STO_CAT", rowD("DIVISION_NAME").ToString.Trim)
                        cmd.Parameters.AddWithValue("@SYS_LUB", Session("usr_nickname"))
                        cmd.Parameters.AddWithValue("@SYS_LUD", System.DateTime.Now)
                        cmd.CommandType = System.Data.CommandType.Text
                        cmd.ExecuteScalar()
                        updateCount = updateCount + 1

                        lblMSG.Text = "STORER data has been already updated"
                    End If

                Next

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted STORER Row : " & insertCount.ToString.Trim & "</p>"
                'strAlert &= "<p>Row Updated : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>STORER has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            Else
                lblMSG.Text = "STORER data has been already updated"
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    'Import PACKING RULE
    Private Sub ImportPACKINGRULEData()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim insertCount As Int16 = 0
        Dim updateCount As Int16 = 0
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            SQLString = "Select * from EBS_WMS_PACKING_RULE"
            Dim codt As New DataTable
            codt = gDB.getDataTable(SQLString, gConn, transaction)
            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then

                'Delete existing customer item rules
                SQLString = "delete from WMS_CUSTOMER_RULE"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                For Each rowD As DataRow In codt.Rows
                    Dim txtInsert As Int16 = 0
                    If rowD("ITEM_ID").ToString.Trim = "" Then
                        If rowD("ACCOUNT_NUMBER").ToString.Trim <> "" Then
                            Dim MSLFlag = 0
                            If rowD("SHELF_DAYS") IsNot Nothing AndAlso rowD("SHELF_DAYS") > 0 Then
                                MSLFlag = 1
                            End If
                            'UPDATE
                            sbCmdText = New StringBuilder()
                            sbCmdText.Append("Update WMS_CUSTOMER Set ")
                            sbCmdText.Append("MIN_SELF_LIFE = @MIN_SELF_LIFE , ")
                            sbCmdText.Append(" MSL_FLAG = @MSL_FLAG")
                            sbCmdText.Append(" Where CUS_CODE=@CUS_CODE And STORER_CODE=@STORER_CODE")
                            cmd = New SqlCommand(sbCmdText.ToString(), gConn, transaction)
                            cmd.Parameters.AddWithValue("@CUS_CODE", rowD("ACCOUNT_NUMBER").ToString.Trim)
                            cmd.Parameters.AddWithValue("@STORER_CODE", rowD("IO_ID").ToString.Trim)
                            cmd.Parameters.AddWithValue("@MIN_SELF_LIFE", rowD("SHELF_DAYS").ToString.Trim)
                            cmd.Parameters.AddWithValue("@MSL_FLAG", MSLFlag)
                            cmd.CommandType = System.Data.CommandType.Text
                            cmd.ExecuteScalar()
                            updateCount = updateCount + 1
                        End If

                    Else
                        SQLString = "Select * from WMS_CUSTOMER_RULE where ITEM_CODE='" & rowD("ITEM_ID").ToString.Trim & "' and CUST_CODE='" & rowD("ACCOUNT_NUMBER").ToString.Trim & "'"
                        Dim codtl As New DataTable
                        codtl = gDB.getDataTable(SQLString, gConn, transaction)
                        If codtl Is Nothing Or codtl.Rows.Count <= 0 Then
                            'INSERT
                            'Parameterized Parameter
                            SQLString = "Insert into WMS_CUSTOMER_RULE(ITEM_CODE,CUST_CODE,MIN_SELF_LIFE,MSL_FLAG,Series) VALUES (@ITEM_CODE,@CUST_CODE,@MIN_SELF_LIFE,@MSL_FLAG,@Series)"
                            cmd = New SqlCommand(SQLString, gConn, transaction)
                            cmd.Parameters.AddWithValue("@ITEM_CODE", rowD("ITEM_ID").ToString.Trim)
                            cmd.Parameters.AddWithValue("@CUST_CODE", rowD("ACCOUNT_NUMBER").ToString.Trim)
                            cmd.Parameters.AddWithValue("@MIN_SELF_LIFE", rowD("SHELF_DAYS").ToString.Trim)
                            cmd.Parameters.AddWithValue("@MSL_FLAG", "1")
                            cmd.Parameters.AddWithValue("@Series", "1")
                            cmd.CommandType = System.Data.CommandType.Text
                            cmd.ExecuteScalar()
                            insertCount = insertCount + 1
                            txtInsert = 1
                        End If
                    End If

                Next

            End If

            transaction.Commit()

            If lblMSG.Text = "" Then
                Dim strAlert As String = "<p>Inserted CUSTOMER RULE Record : " & insertCount.ToString.Trim & "</p>"
                strAlert &= "<p>Updated CUSTOMER Record : " & updateCount.ToString.Trim & "</p>"
                'strAlert &= "<p>STORER has been successfully imported!!</p>"
                lblMSG.Text = strAlert
            End If

        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    Public Sub INSPOSTDATASR(STORER_CODE As String, TRD_BATCH_NO_FR As String, TRD_BATCH_NO_TO As String, TRD_LOC_FR As String, TRD_LOC_TO As String, TR_WH_FR As String, TR_WH_TO As String, ITM_CODE As String, TRD_QTY As Double, TRD_ORG_QTY As Double, TRD_PALLET_NO_FR As String, TRD_PALLET_NO_TO As String, TRD_EXPIRY_DATE As DateTime, transaction As SqlTransaction, gConn As SqlConnection)
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim nextNo As String
        Dim updtSql As String
        Dim tmpSL As String = ""

        Try

            Dim docType As String = ""
            'If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" Then
            '    docType = "SRL"
            'ElseIf Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
            '    docType = "RED"
            'Else
            docType = "STF"
            'End If

            nextNo = DB.getDocNo(docType, gConn, transaction)
            SQLString = "Select * from WMS_STOCK_TRANSFER " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and TR_CODE='" & nextNo & "'"
            Dim dtROCode As DataTable
            dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                'INSERT
                SQLString = "Insert into WMS_STOCK_TRANSFER(IMP_CODE,STORER_CODE,TR_CODE,TR_STATUS,TR_DATE,TR_BY,TR_BATCH_NO,TR_REF_NO,TR_WH_FR,TR_WH_TO,TR_REM,SYS_LUB,SYS_LUD,SYS_CD,SYS_CB,EBS_UPDATE_STATUS) VALUES (@IMP_CODE,@STORER_CODE,@TR_CODE,@TR_STATUS,@TR_DATE,@TR_BY,@TR_BATCH_NO,@TR_REF_NO,@TR_WH_FR,@TR_WH_TO,@TR_REM,@SYS_LUB,@SYS_LUD,@SYS_CD,@SYS_CB,@EBS_UPDATE_STATUS)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@TR_CODE", nextNo)
                cmd.Parameters.AddWithValue("@TR_STATUS", "NEW")
                cmd.Parameters.AddWithValue("@TR_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TR_BY", Session("usr_id"))
                cmd.Parameters.AddWithValue("@TR_BATCH_NO", DBNull.Value)
                cmd.Parameters.AddWithValue("@TR_REF_NO", DBNull.Value)
                cmd.Parameters.AddWithValue("@TR_WH_FR", TR_WH_FR)
                cmd.Parameters.AddWithValue("@TR_WH_TO", TR_WH_TO)
                cmd.Parameters.AddWithValue("@TR_REM", "IPR")
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.Parameters.AddWithValue("@EBS_UPDATE_STATUS", "Y")
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()


                'For ITEM NUMBER
                Dim ITEM_DESCRIPTION As String

                SQLString = "select Top(1) ITEM_DESCRIPTION from EBS_WMS_ITEM_MASTER where IO_ID = '" & STORER_CODE & "' and ITEM_ID='" & ITM_CODE & "'  "
                Dim dtITEM_DESCRIPTION As DataTable
                dtITEM_DESCRIPTION = gDB.getDataTable(SQLString, gConn, transaction)
                If dtITEM_DESCRIPTION IsNot Nothing AndAlso dtITEM_DESCRIPTION.Rows.Count > 0 Then
                    ITEM_DESCRIPTION = dtITEM_DESCRIPTION.Rows(0)("ITEM_DESCRIPTION").ToString.Trim
                Else
                    ITEM_DESCRIPTION = ""
                End If

                'for details'
                SQLString = "Insert into WMS_STOCK_TRANSFER_D(IMP_CODE,STORER_CODE,TR_CODE,ITM_CODE,PACK_KEY,TRD_SEQ,TRD_BATCH_NO,TRD_QTY,TRD_PALLET_NO_FR,TRD_PALLET_NO_TO,TRD_LOC_FR,TRD_LOC_TO,TRD_BATCH_NO_FR,TRD_BATCH_NO_TO,TRD_UOM,TRD_UOM2,TRD_QTY2,TRD_EXPIRY_DATE,TRD_MANU_DATE,TRD_ITM_NAME,TRD_ORG_QTY,TRD_ORG_QTY2,TRD_REQ_QTY,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB) VALUES (@IMP_CODE,@STORER_CODE,@TR_CODE,@ITM_CODE,@PACK_KEY,@TRD_SEQ,@TRD_BATCH_NO,@TRD_QTY,@TRD_PALLET_NO_FR,@TRD_PALLET_NO_TO,@TRD_LOC_FR,@TRD_LOC_TO,@TRD_BATCH_NO_FR,@TRD_BATCH_NO_TO,@TRD_UOM,@TRD_UOM2,@TRD_QTY2,@TRD_EXPIRY_DATE,@TRD_MANU_DATE,@TRD_ITM_NAME,@TRD_ORG_QTY,@TRD_ORG_QTY2,@TRD_REQ_QTY,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@ITM_CODE", ITM_CODE)
                cmd.Parameters.AddWithValue("@TR_CODE", nextNo)
                cmd.Parameters.AddWithValue("@TRD_SEQ", "1")
                cmd.Parameters.AddWithValue("@PACK_KEY", "1")
                cmd.Parameters.AddWithValue("@TRD_BATCH_NO", DBNull.Value)
                cmd.Parameters.AddWithValue("@TRD_QTY", TRD_QTY)
                cmd.Parameters.AddWithValue("@TRD_PALLET_NO_FR", TRD_PALLET_NO_FR)
                cmd.Parameters.AddWithValue("@TRD_PALLET_NO_TO", TRD_PALLET_NO_TO)
                cmd.Parameters.AddWithValue("@TRD_LOC_FR", TRD_LOC_FR)
                cmd.Parameters.AddWithValue("@TRD_LOC_TO", TRD_LOC_TO)
                cmd.Parameters.AddWithValue("@TRD_BATCH_NO_FR", TRD_BATCH_NO_FR)
                cmd.Parameters.AddWithValue("@TRD_BATCH_NO_TO", TRD_BATCH_NO_TO)
                cmd.Parameters.AddWithValue("@TRD_UOM", DBNull.Value)
                cmd.Parameters.AddWithValue("@TRD_UOM2", DBNull.Value)
                cmd.Parameters.AddWithValue("@TRD_QTY2", 0)
                cmd.Parameters.AddWithValue("@TRD_EXPIRY_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TRD_MANU_DATE", DBNull.Value)
                cmd.Parameters.AddWithValue("@TRD_ITM_NAME", ITEM_DESCRIPTION)
                cmd.Parameters.AddWithValue("@TRD_ORG_QTY", TRD_ORG_QTY)
                cmd.Parameters.AddWithValue("@TRD_ORG_QTY2", 0)
                cmd.Parameters.AddWithValue("@TRD_REQ_QTY", 0)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()
            End If

            SQLString = "SELECT * FROM WMS_STOCK_TRANSFER_D " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and TR_CODE='" & nextNo & "'"
            Dim dtRODCode As DataTable
            dtRODCode = gDB.getDataTable(SQLString, gConn, transaction)

            If dtRODCode IsNot Nothing AndAlso dtRODCode.Rows.Count > 0 Then
                For Each rows As DataRow In dtRODCode.Rows
                    'FOR POSTING'
                    st.STORER_CODE = STORER_CODE
                    st.ITM_CODE = gU.decodeNull(rows.Item("ITM_CODE").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("PACK_KEY").ToString.Trim, "")
                    tmpSL = st.getSLInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                    st.IO_CUST_CODE = ""
                    st.IO_AREA = ""
                    st.IO_DOC = "STF"
                    st.IO_DOC_ID = nextNo
                    st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty").ToString.Trim, ""), "0")
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_EXPIRY_DATE = TRD_EXPIRY_DATE.ToString(gU.getConfig("DDFORMAT2"))
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("TRD_MANU_DATE").ToString.Trim, "")

                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("TRD_BATCH_NO_FR").ToString.Trim, "")

                    st.IO_WH = TR_WH_FR
                    st.IO_LOC = gU.decodeNull(rows.Item("TRD_LOC_FR").ToString.Trim.Substring(rows.Item("TRD_LOC_FR").ToString.Trim.Length - 8), "")
                    st.PALLET_NO = gU.decodeNull(rows.Item("TRD_PALLET_NO_FR").ToString.Trim, "")

                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                        st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                        st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty2").ToString.Trim, ""), "0")
                        st.IOS_SL = tmpSL
                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                    End If

                    st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                    st.UpdateStockTrans("OUT", gConn, transaction)
                    st.UpdateStockBalTrans("OUT", gConn, transaction)

                    If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                        st.UpdateStockSerialTrans("OUT", gConn, transaction)
                        st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                    End If

                    If Session("PAGE_SESSION_MENU_CODE") <> "OP_RD" Then
                        st.IO_LOC = gU.decodeNull(rows.Item("trd_loc_to").ToString.Trim.Substring(rows.Item("trd_loc_to").ToString.Trim.Length - 8), "")
                        st.IO_WH = TR_WH_TO

                        If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                            st.IOS_SERIAL_NO = rows.Item("TRD_SERIAL").ToString.Trim
                            st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("trd_qty2").ToString.Trim, ""), "0")
                            st.IOS_SL = tmpSL
                            st.IOS_UOM2 = gU.decodeNull(rows.Item("TRD_UOM2").ToString.Trim, "")
                            Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("TRD_SERIAL").ToString.Trim, ""), gConn, transaction)
                            If gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_TO").ToString.Trim, "")
                            Else
                                st.IOS_DRUM_ID = gU.decodeNull(rows.Item("TRD_DRUM_ID_FR").ToString.Trim, "")
                            End If

                            If gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_TO").ToString.Trim, "")
                            Else
                                st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("TRD_DRUM_LV_FR").ToString.Trim, "")
                            End If
                            st.IOS_REDRUM_YN = "N"
                        End If

                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                        If rows.Item("TRD_SERIAL").ToString.Trim <> "" Then
                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If
                    End If
                Next

            End If

            updtSql = "update wms_stock_transfer " &
                                "Set TR_status = 'POSTED', " &
                                "sys_lub = '" & Session("usr_id") & "', " &
                                "sys_lud = Getdate() " &
                                "where imp_code = 'WMS' " &
                                "and storer_code = '" & STORER_CODE & "' " &
                                "and tr_code = '" & nextNo & "' "

            gDB.amendData(updtSql, gConn, transaction)

        Catch ex As Exception
            WriteExceptionLog(ex)
            Throw ex
        End Try
    End Sub

    Public Sub SendMailMatchingOutRpt()
        Dim gConn = gDB.getConnection()
        Dim sqlString As String
        Dim nDataSource As DataSet


        Try
            'For OUTSTANDING REPORT'
            sqlString = " exec sp_MachingOutRpt "

            nDataSource = gDB.getDataSet(sqlString, gConn)

            Dim excelFolder As String
            excelFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempFiles\" & "_MatchingOutstandingReportItems_" & System.DateTime.UtcNow.ToString("MM-dd-yyyy hh-mm-ss") & ".xls")
            Dim _Result As Boolean = WriteXLSFile(excelFolder, nDataSource)
            If _Result = True Then
                Dim Smtp_Server As New SmtpClient
                Dim e_mail As New MailMessage()
                'Smtp_Server.UseDefaultCredentials = True
                'Smtp_Server.Host = "hksmtp.lamsoon.com"
                'e_mail = New MailMessage()
                'e_mail.From = New MailAddress("hkwms@lamsoon.com")
                'e_mail.To.Add("wmsexceptionreport@lamsoon.com")
                'e_mail.Subject = "WMS Exception Report from WMS on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
                'e_mail.IsBodyHtml = False
                'e_mail.Body = "Kindly find the report attachment."
                'e_mail.Attachments.Add(New Attachment(excelFolder))
                'Smtp_Server.Send(e_mail)

                Smtp_Server.UseDefaultCredentials = False
                Smtp_Server.Credentials = New Net.NetworkCredential("augursmail@gmail.com", "Augurs@0009")
                Smtp_Server.Port = 587
                Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "smtp.gmail.com"

                e_mail = New MailMessage()
                e_mail.From = New MailAddress("augursmail@gmail.com")
                e_mail.To.Add("madhvendra009@gmail.com")
                'e_mail.To.Add("robertwong68@gmail.com")
                e_mail.Subject = "WMS Exception Report from WMS on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
                e_mail.IsBodyHtml = False
                e_mail.Body = "Kindly find the report attachment."
                e_mail.Attachments.Add(New Attachment(excelFolder))
                Smtp_Server.Send(e_mail)
            End If
        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
        End Try
    End Sub

    Private Sub SendMailToUser(excelFolder As String, ToEmail As String, IO_CODE As String, BatchNo As String, EBSUser As String)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()

            Dim EmailIDs = gU.getConfig("shortage_report_receiving_email_address")
            If EmailIDs IsNot Nothing Then
                Dim EmailList = EmailIDs.Split(",")
                For Each EmailID As String In EmailList
                    e_mail.To.Add(EmailID)
                Next
            End If

            'Smtp_Server.UseDefaultCredentials = True
            'Smtp_Server.Host = "hksmtp.lamsoon.com"
            ''e_mail = New MailMessage()
            'e_mail.From = New MailAddress("hkwms@lamsoon.com")
            'e_mail.To.Add(ToEmail)
            'e_mail.Subject = "Sales Admin Shortage report from WMS for IO Code '" & IO_CODE & "' and SO Batch No: " & BatchNo & " and EBS User: " & EBSUser & " on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
            'e_mail.IsBodyHtml = False
            'e_mail.Body = "Kindly find the report attachment."
            'e_mail.Attachments.Add(New Attachment(excelFolder))
            'Smtp_Server.Send(e_mail)

            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("augursmail@gmail.com", "Augurs@0009")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            'e_mail = New MailMessage()
            e_mail.From = New MailAddress("augursmail@gmail.com")
            e_mail.To.Add("madhvendra009@gmail.com")
            'e_mail.To.Add("robertwong68@gmail.com")
            e_mail.Subject = "Sales Admin Shortage report from WMS for IO Code '" & IO_CODE & "' and SO Batch No: " & BatchNo & " and EBS User: " & EBSUser & " on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
            e_mail.IsBodyHtml = False
            e_mail.Body = "Kindly find the report attachment."
            e_mail.Attachments.Add(New Attachment(excelFolder))
            Smtp_Server.Send(e_mail)

        Catch error_t As Exception
            lblMSG.Text = error_t.ToString
        End Try
    End Sub

    Public Function WriteXLSFile(ByVal pFileName As String, ByVal pDataSet As DataSet) As Boolean
        Try
            'Create a workbook instance
            Dim workbook As Workbook = New Workbook()
            Dim worksheet As Worksheet
            Dim iRow As Integer = 0
            Dim iCol As Integer = 0
            Dim sTemp As String = String.Empty
            Dim dTemp As Double = 0
            Dim iTemp As Integer = 0
            Dim dtTemp As DateTime
            Dim count As Integer = 0
            Dim iTotalRows As Integer = 0
            Dim iSheetCount As Integer = 0

            'Read DataSet
            If Not pDataSet Is Nothing And pDataSet.Tables.Count > 0 Then

                Dim ErrorReport As String = "N"
                If (pDataSet.Tables.Count = 5) Then
                    ErrorReport = "Y"
                End If
                'Traverse DataTable inside the DataSet
                For Each dt As DataTable In pDataSet.Tables

                    'Create a worksheet instance
                    iSheetCount = iSheetCount + 1
                    If (ErrorReport = "Y") Then
                        If (iSheetCount = 1) Then
                            worksheet = New Worksheet("IPR Outstanding Report")
                        ElseIf (iSheetCount = 2) Then
                            worksheet = New Worksheet("Non PO")
                        ElseIf (iSheetCount = 3) Then
                            worksheet = New Worksheet("Stock Return")
                        ElseIf (iSheetCount = 4) Then
                            worksheet = New Worksheet("Stock Adjustment")
                        Else
                            worksheet = New Worksheet("Compare Stock Balance")
                        End If
                    Else
                        If (iSheetCount = 1) Then
                            worksheet = New Worksheet("Sales Admin Shortage Report")
                        ElseIf (iSheetCount = 2) Then
                            worksheet = New Worksheet("Item Balance Report")
                        ElseIf (iSheetCount = 3) Then
                            worksheet = New Worksheet("Non allocated Item Balance Report")
                        Else
                            worksheet = New Worksheet("NEW NonAllocated report")
                        End If
                    End If

                    'Write Table Header
                    iCol = 0
                    For Each dc As DataColumn In dt.Columns
                        worksheet.Cells(0, iCol) = New Cell(dc.ColumnName)
                        iCol = iCol + 1
                    Next

                    'Write Table Body
                    iRow = 1
                    For Each dr As DataRow In dt.Rows
                        iCol = 0
                        For Each dc As DataColumn In dt.Columns
                            sTemp = dr(dc.ColumnName).ToString()
                            Select Case dc.DataType
                                Case GetType(DateTime)
                                    DateTime.TryParse(sTemp, dtTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dtTemp, "MM/DD/YYYY")
                                Case GetType(Double)
                                    Double.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                                Case GetType(Decimal)
                                    Decimal.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                                Case Else
                                    If Int32.TryParse(sTemp, iTemp) Then
                                        worksheet.Cells(iRow, iCol) = New Cell(Convert.ToInt32(iTemp), "0")
                                    Else
                                        worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                                    End If
                            End Select
                            iCol = iCol + 1
                        Next
                        iRow = iRow + 1
                    Next

                    'Attach worksheet to workbook
                    workbook.Worksheets.Add(worksheet)
                    iTotalRows = iTotalRows + iRow
                Next
            End If

            'Bug on Excel Library, min file size must be 7 Kb
            'thus we need to add empty row for safety
            If iTotalRows < 100 Then
                worksheet = New Worksheet("Sheet X")
                count = 1
                Do While count < 100
                    worksheet.Cells(count, 0) = New Cell(" ")
                    count = count + 1
                Loop
                workbook.Worksheets.Add(worksheet)
            End If

            workbook.Save(pFileName)
            Return True
        Catch ex As Exception
            WriteExceptionLog(ex)
            Return False
        End Try
    End Function

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub

    Private Sub SendEmail(gConn As SqlConnection, transaction As SqlTransaction, Storer1 As String, Storer2 As String, BatchNo As String, EBSUser As String)

        Dim sqlString As String
        Dim nDataSource As DataSet

        nDataSource = gDB.getDataSet("exec sp_StorageBalReport " & Storer1 & "," & Storer2, gConn, transaction)

        Dim excelFolder As String
        excelFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempFiles\" & "_SalesAdminReportItems_" & System.DateTime.UtcNow.ToString("MM-dd-yyyy hh-mm-ss") & ".xls")
        Dim _Result As Boolean = WriteXLSFile(excelFolder, nDataSource)
        If _Result = True Then
            If (Storer1 = "104") Then
                SendMailToUser(excelFolder, "terence.chan@lamsoon.com", "201 and 204", BatchNo, EBSUser)
            ElseIf (Storer1 = "105") Then
                SendMailToUser(excelFolder, "horace.chan@lamsoon.com", "301", BatchNo, EBSUser)
            ElseIf (Storer1 = "106") Then
                SendMailToUser(excelFolder, "ken.sin@lamsoon.com", "401", BatchNo, EBSUser)
            End If
        End If
    End Sub

    Private Sub CloseSQLConnection()
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim gConn = gDB.getConnection()
        lblMSG.Text = ""
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try
            sbCmdText = New StringBuilder()
            sbCmdText.Append("USE	MASTER GO DECLARE @kill varchar(8000) = ''; SELECT @kill = @kill + 'kill ' + CONVERT(varchar(5), session_id) + ';'   ")
            sbCmdText.Append("FROM sys.dm_exec_sessions WHERE database_id  = db_id('prod_debug') exec(@kill) ")
            cmd = New SqlCommand(sbCmdText.ToString(), gConn)
        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

    End Sub

    Private Sub DelTempDir()
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim DelSOSQL As String

        Try

            Dim ExcelFiles As String() = Directory.GetFiles(System.Configuration.ConfigurationManager.AppSettings.Item("tmpExcelFiles"))
            For Each file As String In ExcelFiles
                Dim fi As FileInfo = New FileInfo(file)
                If fi.LastAccessTime < DateTime.Now.AddDays(-7) Then
                    fi.Delete()
                End If
            Next

            Dim PDFFiles As String() = Directory.GetFiles(System.Configuration.ConfigurationManager.AppSettings.Item("tmpPDFFiles"))
            For Each file As String In PDFFiles
                Dim fi As FileInfo = New FileInfo(file)
                If fi.LastAccessTime < DateTime.Now.AddDays(-7) Then
                    fi.Delete()
                End If
            Next

            ''Write Action Log data to file
            Dim SQLString As String = "Select * from ActionLogs where TransDate < GetDate()-7 "
            Dim ActionDt As New DataTable
            ActionDt = gDB.getDataTable(SQLString, gConn, transaction)

            Dim ActionLogFile As String = Server.MapPath("~/Logs/ActionLog " + DateTime.Now.ToString("dd-MM-yyyy hh-mm-ss tt") + ".txt")
            Dim logFile As FileStream = File.Create(ActionLogFile)
            logFile.Close()
            Dim message As StringBuilder = New StringBuilder()
            message.Append(String.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            message.Append(Environment.NewLine)
            For Each rowD As DataRow In ActionDt.Rows
                message.Append(String.Format("Record: {0}", rowD("ID").ToString.Trim + ", " + rowD("ActionDetails").ToString.Trim + ", " + rowD("ActionType").ToString.Trim + ", " + rowD("TableName").ToString.Trim + ", " + rowD("TransDate").ToString.Trim + ", " + rowD("ActionUserId").ToString.Trim))
                message.Append(Environment.NewLine)
            Next
            Using writer As New StreamWriter(ActionLogFile.ToString(), True)
                writer.WriteLine(message)
                writer.Close()
            End Using

            ''Delete SO Trail old data and Log data
            DelSOSQL = " delete from EBS_WMS_TRANS_ITX_SO_DETAIL where TRANSACTION_ID in (select TRANSACTION_ID from EBS_WMS_TRANS_ITX_SO_HEADER where TRIAL_TYPE='TRIAL' and CREATION_DATE < GetDate()-7)  " &
                   " delete from EBS_WMS_TRANS_ITX_ACTION where TRANSACTION_ID in (select TRANSACTION_ID from EBS_WMS_TRANS_ITX_SO_HEADER where TRIAL_TYPE='TRIAL' and CREATION_DATE < GetDate()-7) " &
                   " delete from EBS_WMS_TRANS_ITX_SO_HEADER where TRIAL_TYPE='TRIAL' and CREATION_DATE < GetDate()-7 " &
                   " delete from ActionLogs where TransDate < GetDate()-7 "
            gDB.amendData(DelSOSQL, gConn, transaction)

            transaction.Commit()

        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Sub ArchiveData()
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim ArchiveDataSQL As String

        Try
            ''Archive data SQL
            ArchiveDataSQL = " exec SP_ArchiveData " + System.Configuration.ConfigurationManager.AppSettings.Item("KEEPDATADAYS")
            gDB.amendData(ArchiveDataSQL, gConn, transaction)
            transaction.Commit()

        Catch ex As Exception
            WriteExceptionLog(ex)
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Sub WriteExceptionLog(ex As Exception)
        Dim message As String = String.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        message += Environment.NewLine
        message += "-----------------------------------------------------------"
        message += Environment.NewLine
        message += String.Format("Message: {0}", ex.Message)
        message += Environment.NewLine
        message += String.Format("StackTrace: {0}", ex.StackTrace)
        message += Environment.NewLine
        message += String.Format("Source: {0}", ex.Source)
        message += Environment.NewLine
        message += String.Format("TargetSite: {0}", ex.TargetSite.ToString())
        message += Environment.NewLine
        message += "-----------------------------------------------------------"
        message += Environment.NewLine
        Dim path As String = Server.MapPath("~/ErrorLog.txt")
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(message)
            writer.Close()
        End Using
    End Sub

    Public Sub SendWrongLOTEmail()
        Dim gConn = gDB.getConnection()
        Dim sqlString As String
        Dim nDataSource As DataTable


        Try
            sqlString = " SELECT	STORER_CODE,ITM_CODE,ILOC_WH,ILOC_BATCH_NO,ILOC_LOC, ISDATE(LEFT(ILOC_BATCH_NO, 4)+ '-' + RIGHT(LEFT(ILOC_BATCH_NO, 6),2)+ " &
            " '-' + RIGHT(ILOC_BATCH_NO, 2)) AS Result, ILOC_BATCH_NO FROM [dbo].[WMS_ITEM_LOC_BAL] WHERE	LEN(ILOC_BATCH_NO)<>  8 or  " &
            " IsDate(	LEFT(ILOC_BATCH_NO, 4)+ '-' + RIGHT(LEFT(ILOC_BATCH_NO, 6),2)+ '-' + RIGHT(ILOC_BATCH_NO, 2)) <>1 ORDER BY	Result "

            nDataSource = gDB.getDataTable(sqlString, gConn)

            If nDataSource.Rows.Count > 0 Then
                Dim MailBody As String
                MailBody = "<p>The below list is of wrong LOT fromat in WMS_ITEM_LOC_BAL table which may cause issue in Lot Allocation Program. Please update in correct format immediately.</p>"
                MailBody += "<table>"
                MailBody += "<tr><th style='border: 1px solid;'>Storer Code</th><th style='border: 1px solid;'>Item Code</th><th style='border: 1px solid;'>Warehouse</th><th style='border: 1px solid;'>Location</th><th style='border: 1px solid;'>Batch No</th></tr>"
                For Each rowD As DataRow In nDataSource.Rows
                    MailBody += "<tr><td style='border: 1px solid;'>" + rowD("STORER_CODE").ToString.Trim + "</td><td style='border: 1px solid;'>" + rowD("ITM_CODE").ToString.Trim + "</td><td style='border: 1px solid;'>" + rowD("ILOC_WH").ToString.Trim + "</td><td style='border: 1px solid;'>" + rowD("ILOC_LOC").ToString.Trim + "</td><td style='border: 1px solid;'>" + rowD("ILOC_BATCH_NO").ToString.Trim + "</td></tr>"
                Next
                MailBody += "</table>"

                Dim Smtp_Server As New SmtpClient
                Dim e_mail As New MailMessage()

                'Smtp_Server.UseDefaultCredentials = True
                'Smtp_Server.Host = "hksmtp.lamsoon.com"
                'e_mail = New MailMessage()
                'e_mail.From = New MailAddress("hkwms@lamsoon.com")
                'e_mail.To.Add("wmsexceptionreport@lamsoon.com")
                'e_mail.Subject = "Urgent Action Needed. WMS Wrong LOT format in Stock Balance on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
                'e_mail.IsBodyHtml = True
                'e_mail.Body = MailBody
                'Smtp_Server.Send(e_mail)

                Smtp_Server.UseDefaultCredentials = False
                Smtp_Server.Credentials = New Net.NetworkCredential("augursmail@gmail.com", "Augurs@0009")
                Smtp_Server.Port = 587
                Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "smtp.gmail.com"

                e_mail = New MailMessage()
                e_mail.From = New MailAddress("augursmail@gmail.com")
                e_mail.To.Add("madhvendra009@gmail.com")
                e_mail.To.Add("robertwong68@gmail.com")
                e_mail.Subject = "Urgent Action Needed. WMS Wrong LOT format in Stock Balance on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
                e_mail.IsBodyHtml = True
                e_mail.Body = MailBody
                Smtp_Server.Send(e_mail)
            End If
        Catch ex As Exception
            WriteExceptionLog(ex)
            Response.Write(ex.Message)
        End Try
    End Sub

    Private Sub ReIndexing()
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim ReIndexDataSQL As String

        Try
            ReIndexDataSQL = "  ALTER INDEX PK_EBS_WMS_TRANS_ITX_ACTION1 ON EBS_WMS_TRANS_ITX_ACTION REBUILD " &
             " ALTER Index PK_EBS_WMS_TRANS_ITX_SO_HEADER ON EBS_WMS_TRANS_ITX_SO_HEADER REBUILD " &
             " ALTER INDEX PK_WMS_CUST_ORDER_D ON WMS_CUST_ORDER_D REBUILD " &
             " ALTER Index PK_WMS_CUST_ORDER ON WMS_CUST_ORDER REBUILD " &
             " ALTER INDEX PK_WMS_DELV_ORDER_D ON WMS_DELV_ORDER_D REBUILD " &
             " ALTER Index PK_WMS_DELV_ORDER ON WMS_DELV_ORDER REBUILD " &
             " ALTER INDEX PK_WMS_DELV_ORDER_TEMP_1 ON WMS_DELV_ORDER_TEMP REBUILD " &
             " ALTER Index PK_WMS_DO_PICKLIST_D ON WMS_DO_PICKLIST_D REBUILD " &
             " ALTER INDEX PK_EBS_WMS_TRANS_ITX_ACTION ON WMS_EBS_TRANS_ITX_ACTION REBUILD " &
             " ALTER Index PK_WMS_EBS_TRANS_ITX_SO ON WMS_EBS_TRANS_ITX_SO REBUILD " &
             " ALTER INDEX PK_WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT1 ON WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT REBUILD " &
             " ALTER Index PK__WMS_IN_T__6CD7C2F5C4774E4E ON WMS_IN_TX REBUILD " &
             " ALTER INDEX PK_WMS_ITEM_LOC_BAL_TX ON WMS_ITEM_LOC_BAL_TX  REBUILD " &
             " ALTER Index [SQLPROPK_dbo.WMS_OUT_TX_20210116] ON [dbo].[WMS_OUT_TX] REBUILD " &
             " ALTER INDEX PK_EBS_WMS_TRANS_ITX_IPR_DETAIL ON EBS_WMS_TRANS_ITX_IPR_DETAIL REBUILD " &
             " ALTER Index PK_EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT ON EBS_WMS_TRANS_ITX_IPR_DETAIL_LOT REBUILD " &
             " ALTER INDEX PK_EBS_WMS_TRANS_ITX_STOCK_ADJUSTMENT ON EBS_WMS_TRANS_ITX_STOCK_ADJUSTMENT REBUILD " &
             " ALTER Index PK_WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER ON WMS_EBS_TRANS_ITX_INVENTORY_TRANSFER REBUILD " &
             " ALTER INDEX PK_WMS_EBS_TRANS_ITX_IPR_GRN ON WMS_EBS_TRANS_ITX_IPR_GRN REBUILD "
            gDB.amendData(ReIndexDataSQL, gConn, transaction)
            transaction.Commit()

        Catch ex As Exception
            WriteExceptionLog(ex)
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Sub GenerateBackup()
        Dim gConn = gDB.getConnection()
        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Dim BackupDataSQL As String

        Try
            ''Backup data SQL
            Dim BakPath As String = "D:\\LSDBBackup\\" + System.DateTime.Now.Date.Day.ToString() + System.DateTime.Now.Date.Month.ToString() + System.DateTime.Now.Date.Year.ToString() + ".bak"
            BackupDataSQL = "backup database [LAMSOONITXDB] to disk =" + "'" + BakPath + "'"
            gDB.amendData(BackupDataSQL, gConn, transaction)
            transaction.Commit()

        Catch ex As Exception
            WriteExceptionLog(ex)
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub
End Class
