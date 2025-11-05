Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq
Partial Class CycleCount
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private lstError As New List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then
            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
        End If
    End Sub

    Protected Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim gconn = gDB.getConnection()
        Dim tran As SqlTransaction
        tran = gconn.BeginTransaction
        Try

            Dim A_Per = 0.1
            Dim B_Per = 0.2
            Dim ITM_TOTAL As Int16 = 0
            If STORER_CODE.SelectedValue IsNot Nothing AndAlso STORER_CODE.SelectedValue.ToString() <> "" Then

                Dim SQLString = "Select IsNull(Count(cc.Itm_code),0)ItmCount from WMS_ITEM cc where cc.STORER_CODE = '" + STORER_CODE.SelectedValue + "' and (Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE and stk.ILOC_WH ='FG01') > 0"
                Dim dtTemp = gDB.getDataTable(SQLString, gconn, tran)
                If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ITM_TOTAL = Convert.ToInt32(dtTemp.Rows(0)("ItmCount"))

                    SQLString = "Update WMS_ITEM Set ITM_MFG = 'C' where STORER_CODE = '" + STORER_CODE.SelectedValue + "'"
                    gDB.amendData(SQLString, gconn, tran)

                    SQLString = "Select  SUM(cc.IO_QTY)Qty,cc.ITM_CODE,Count(cc.ITM_CODE)ITM_Count from WMS_OUT_TX cc Where cc.IO_DOC = 'DO' and (getdate() <= DATEADD(MONTH,3, cc.IO_DATETIME)) and (Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE and stk.ILOC_WH ='FG01' ) > 0 and cc.STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "'	GROUP BY cc.ITM_CODE	ORDER BY SUM(cc.IO_QTY) DESC OFFSET 0 ROWS  FETCH NEXT " + (Convert.ToInt32(A_Per * ITM_TOTAL)).ToString + " ROWS ONLY"
                    Dim dtOUTA = gDB.getDataTable(SQLString, gconn, tran)
                    Dim A_ITEMS As String = ""
                    For Each row As DataRow In dtOUTA.Rows
                        SQLString = "Update WMS_ITEM Set ITM_MFG = 'A' where ITM_CODE = '" + row("ITM_CODE").ToString() + "' and STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "'"
                        gDB.amendData(SQLString, gconn, tran)
                        'If A_ITEMS = "" Then
                        '    A_ITEMS = "'" + row("ITM_CODE").ToString() + "'"
                        'Else
                        '    A_ITEMS = A_ITEMS + ",'" + row("ITM_CODE").ToString() + "'"
                        'End If
                    Next
                    'If A_ITEMS <> "" Then
                    '    SQLString = "Select Top " + (Int((A_Per + B_Per) * ITM_TOTAL)).ToString + " SUM(IO_QTY)Qty,ITM_CODE,Count(ITM_CODE)ITM_Count from WMS_OUT_TX Where IO_DOC = 'DO' and (IO_DATETIME <= DATEADD(MONTH,3, getdate()) and IO_DATETIME >= getdate()) and STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' and ITM_CODE not in (" + A_ITEMS + ")	GROUP BY ITM_CODE ORDER BY SUM(IO_QTY) DESC"
                    'Else
                    '    SQLString = "Select Top " + (Int((A_Per + B_Per) * ITM_TOTAL)).ToString + " SUM(IO_QTY)Qty,ITM_CODE,Count(ITM_CODE)ITM_Count from WMS_OUT_TX Where IO_DOC = 'DO' and (IO_DATETIME <= DATEADD(MONTH,3, getdate()) and IO_DATETIME >= getdate()) and STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' GROUP BY ITM_CODE	ORDER BY SUM(IO_QTY) DESC"
                    'End If
                    SQLString = "Select SUM(cc.IO_QTY)Qty,cc.ITM_CODE,Count(cc.ITM_CODE)ITM_Count from WMS_OUT_TX cc inner join wms_item itm on cc.ITM_CODE=itm.Itm_Code  AND CC.STORER_CODE =ITM.STORER_CODE Where IO_DOC = 'DO' And itm.ITM_MFG != 'A' and (getdate() <= DATEADD(MONTH,3, cc.IO_DATETIME)) and (Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE and stk.ILOC_WH ='FG01' ) > 0 and cc.STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "'	GROUP BY cc.ITM_CODE	ORDER BY Sum(cc.IO_QTY) asc OFFSET " + (Convert.ToInt32((A_Per) * ITM_TOTAL)).ToString + "  ROWS  FETCH NEXT " + (Convert.ToInt32((B_Per) * ITM_TOTAL)).ToString + " ROWS ONLY"
                    Dim dtOUTB = gDB.getDataTable(SQLString, gconn, tran)
                    For Each row As DataRow In dtOUTB.Rows
                        SQLString = "Update WMS_ITEM Set ITM_MFG = 'B' where ITM_CODE = '" + row("ITM_CODE").ToString() + "' and STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' AND ITM_MFG != 'A'"
                        gDB.amendData(SQLString, gconn, tran)
                    Next

                    SQLString = "Delete from CC_ITEM Where STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' And WH_LOC='' And WH_CODE='' And LOT_NO=''"
                    gDB.amendData(SQLString, gconn, tran)
                    SQLString = "Select a.*  from WMS_ITEM a Where a.STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' and (Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where a.IMP_CODE=stk.IMP_CODE and a.ITM_CODE=stk.ITM_CODE and a.STORER_CODE=stk.STORER_CODE and stk.ILOC_WH ='FG01' ) > 0 ORDER BY ITM_MFG"
                    Dim dtItem = gDB.getDataTable(SQLString, gconn, tran)
                    If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then

                        Dim SEQ_No As Int16 = 1
                        For Each row As DataRow In dtItem.Rows
                            SQLString = "Insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDATE,CBYFK) values('" + row("IMP_CODE") + "','" + row("STORER_CODE") + "','" + row("ITM_CODE") + "','" + (SEQ_No).ToString() + "','NEW','" + System.DateTime.Now + "','LS_ADMIN')"
                            gDB.amendData(SQLString, gconn, tran)
                            SEQ_No = SEQ_No + 1
                        Next
                    End If
                    Dim A_POS_1 As Double = Int(ITM_TOTAL / 3)
                    Dim A_POS_2 As Double = Int(ITM_TOTAL / 3 * 2)
                    Dim B_POS_1 As Double = Int(ITM_TOTAL / 2)

                    SQLString = "Select *  from WMS_ITEM Where STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' AND ITM_MFG = 'A'"
                    dtItem = gDB.getDataTable(SQLString, gconn, tran)

                    For Each row As DataRow In dtItem.Rows
                        Dim dml_SEQ_No As Double = A_POS_1 + Convert.ToDouble("0.001")
                        SQLString = "Insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDATE,CBYFK) values('" + row("IMP_CODE") + "','" + row("STORER_CODE") + "','" + row("ITM_CODE") + "','" + Math.Round(dml_SEQ_No, 3).ToString() + "','NEW','" + System.DateTime.Now + "','LS_ADMIN')"
                        gDB.amendData(SQLString, gconn, tran)
                        A_POS_1 = dml_SEQ_No
                    Next
                    SQLString = "Select *  from WMS_ITEM Where STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' AND ITM_MFG = 'A'"
                    dtItem = gDB.getDataTable(SQLString, gconn, tran)

                    For Each row As DataRow In dtItem.Rows
                        Dim dml_SEQ_No As Double = A_POS_2 + Convert.ToDouble("0.001")
                        SQLString = "Insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDATE,CBYFK) values('" + row("IMP_CODE") + "','" + row("STORER_CODE") + "','" + row("ITM_CODE") + "','" + Math.Round(dml_SEQ_No, 3).ToString() + "','NEW','" + System.DateTime.Now + "','LS_ADMIN')"
                        gDB.amendData(SQLString, gconn, tran)
                        A_POS_2 = dml_SEQ_No
                    Next
                    SQLString = "Select * from WMS_ITEM Where STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "' AND ITM_MFG = 'B'"
                    dtItem = gDB.getDataTable(SQLString, gconn, tran)

                    For Each row As DataRow In dtItem.Rows
                        Dim dml_SEQ_No As Double = B_POS_1 + Convert.ToDouble("0.001")
                        SQLString = "Insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDATE,CBYFK) values('" + row("IMP_CODE") + "','" + row("STORER_CODE") + "','" + row("ITM_CODE") + "','" + Math.Round(dml_SEQ_No, 3).ToString() + "','NEW','" + System.DateTime.Now + "','LS_ADMIN')"
                        gDB.amendData(SQLString, gconn, tran)
                        B_POS_1 = dml_SEQ_No

                    Next
                    tran.Commit()
                    uiFun.displayMsg(Me, "", "Cycle Count Success.", Session("gLang"))
                    lblMsg.Text = "Cycle Count Success."
                End If
            End If
        Catch ex As Exception
            tran.Rollback()
            lblMsg.Text = ex.Message
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        Finally
            getCCItems()
        End Try
    End Sub



    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles STORER_CODE.SelectedIndexChanged
        Dim SQLString = "Select ISNULL(STO_CC,'1')STO_CC from WMS_STORER Where IMP_CODE='WMS' and STORER_CODE = '" + STORER_CODE.SelectedValue.ToString + "'"
        Dim dtSTO = gDB.getDataTable(SQLString)
        If dtSTO IsNot Nothing And dtSTO.Rows.Count > 0 Then
            txtPickCount.Text = dtSTO.Rows(0)("STO_CC")
        End If
    End Sub

    Protected Sub btnDailyCount_Click(sender As Object, e As EventArgs)
        Dim gconn = gDB.getConnection()
        Dim tran As SqlTransaction
        tran = gconn.BeginTransaction
        Try

            Dim A_Per = 10
            Dim B_Per = 20
            Dim ITM_TOTAL As Int16 = 0
            If STORER_CODE.SelectedValue IsNot Nothing AndAlso STORER_CODE.SelectedValue.ToString() <> "" Then
                If txtPickCount.Text IsNot Nothing AndAlso txtPickCount.Text.ToString() <> "" Then

                    Dim SQLString As String = "Update CC_ITEM set STATUS='NO_STOCK' where Item_code in (select itm_code from WMS_ITEM_LOC_BAL  where STORER_CODE='" + STORER_CODE.SelectedValue.ToString.Trim + "' and IMP_CODE='WMS' group by itm_code having sum(ILOC_BAL_QTY)<=0.0) and STORER_CODE='" + STORER_CODE.SelectedValue.ToString.Trim + "' and IMP_CODE='WMS' "
                    gDB.amendData(SQLString, gconn, tran)

                    Dim PickCount = Convert.ToInt32(txtPickCount.Text.Trim)
                    'Dim SQLString As String = "Select Top " + PickCount.ToString + " cc.CC_ID,cc.IMP_CODE,cc.STORER_CODE,cc.ITEM_CODE,cc.SEQ_NO,cc.STATUS,cc.CDATE,cc.CByFk from CC_ITEM cc Where cc.STORER_CODE = '" + STORER_CODE.SelectedValue.ToString.Trim + "' And STATUS='NEW' And ((Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITEM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE ) > 0 Or ISNULL(cc.LOT_NO,'') != '') ORDER BY  cc.LOT_NO desc, cc.SEQ_NO ASC "
                    SQLString = "Select   Top " + PickCount.ToString + " cc.ITEM_CODE from CC_ITEM cc Where cc.STORER_CODE = '" + STORER_CODE.SelectedValue.ToString.Trim + "' And STATUS='NEW' And ((Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITEM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE ) > 0 Or ISNULL(cc.LOT_NO,'') != '') Group by cc.ITEM_CODE ORDER BY  MAX(cc.LOT_NO) desc, MAX(cc.SEQ_NO) ASC "
                    Dim dtTemp = gDB.getDataTable(SQLString, gconn, tran)
                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then

                        Dim nextNo = DB.getDocNo("CHK")
                        SQLString = "Insert into WMS_STOCK_CHECK (IMP_CODE,STORER_CODE,CK_CODE,CK_STATUS,CK_TYPE,CK_DATE,SYS_CD,CK_WH) values('WMS','" + STORER_CODE.SelectedValue.ToString.Trim + "','" + nextNo + "','NEW','DAILY','" + System.DateTime.Now + "','" + System.DateTime.Now + "','FG01')"
                        gDB.amendData(SQLString, gconn, tran)
                        Dim SEQ_No As Int16 = 0
                        For Each row As DataRow In dtTemp.Rows

                            SQLString = "select * from  WMS_ITEM_LOC_BAL where ITM_CODE = '" + row("ITEM_CODE").ToString() + "' and (ILOC_BAL_QTY> 0)  and ILOC_WH ='FG01' and STORER_CODE='" + STORER_CODE.SelectedValue.ToString.Trim + "' "
                            Dim dtStk = gDB.getDataTable(SQLString, gconn, tran)
                            If dtStk IsNot Nothing AndAlso dtStk.Rows.Count > 0 Then

                                For Each rowD As DataRow In dtStk.Rows
                                    SQLString = "Insert into WMS_STOCK_CHECK_D (IMP_CODE,STORER_CODE,CK_CODE,CKD_SEQ,CKD_ITM_CODE,CKD_PACK_KEY,CKD_LOC,CKD_ORG_QTY,CKD_BATCH_NO,CKD_STATUS,CKD_PALLET_NO,CKD_EXPIRY_DATE,SYS_CD, CKD_DRUM_ID) values('" + rowD("IMP_CODE").ToString + "','" + rowD("STORER_CODE").ToString + "','" + nextNo.ToString + "','" + (SEQ_No + 1).ToString + "','" + rowD("ITM_CODE").ToString + "','" + rowD("PACK_KEY").ToString + "','" + rowD("ILOC_LOC").ToString + "','" + rowD("ILOC_BAL_QTY").ToString + "','" + rowD("ILOC_BATCH_NO").ToString + "','NEW','" + rowD("ILOC_PALLET_NO").ToString + "','" + rowD("ILOC_EXPIRY_DATE").ToString + "','" + System.DateTime.Now + "','DAILY')"
                                    gDB.amendData(SQLString, gconn, tran)
                                    SEQ_No = SEQ_No + 1
                                Next

                            End If
                            SQLString = "Update CC_ITEM Set STATUS='COMPLETED' Where ITEM_CODE='" + row("ITEM_CODE").ToString() + "'"
                            gDB.amendData(SQLString, gconn, tran)
                        Next
                        tran.Commit()
                        uiFun.displayMsg(Me, "", "Daily Count Success.", Session("gLang"))
                        lblMsg.Text = "Daily Count Success."
                    End If

                Else
                    uiFun.displayMsg(Me, "", "Enter Pick Count", Session("gLang"))
                End If
            End If
        Catch ex As Exception
            tran.Rollback()
            lblMsg.Text = ex.Message
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))

        Finally
            getCCItems()
        End Try
    End Sub

    Private Sub getCCItems()
        If STORER_CODE.SelectedValue IsNot Nothing AndAlso STORER_CODE.SelectedValue.ToString() <> "" Then
            Dim SQLString = "Select cc.SEQ_NO,cc.STORER_CODE,cc.ITEM_CODE,itm.ITM_SKU_NO,itm.ITM_DESC,itm.ITM_MFG,(Select IsNUll(Sum(stk.ILOC_BAL_QTY),0)as QTY from WMS_ITEM_LOC_BAL stk Where cc.IMP_CODE=stk.IMP_CODE and cc.ITEM_CODE=stk.ITM_CODE and cc.STORER_CODE=stk.STORER_CODE )ITM_QTY,cc.WH_CODE,cc.LOT_NO,cc.STATUS,cc.CDATE,cc.CByFk from CC_ITEM cc Inner Join WMS_ITEM itm on cc.ITEM_CODE=itm.ITM_CODE and cc.IMP_CODE = itm.IMP_CODE and cc.STORER_CODE = itm.STORER_CODE Where cc.STATUS='NEW' and cc.STORER_CODE = '" + STORER_CODE.SelectedValue + "' ORDER BY  cc.LOT_NO desc, cc.SEQ_NO ASC "
            Dim dtTemp = gDB.getDataTable(SQLString)
            Dim ttlnew As Int16 = 0
            Dim ttlnewnostk As Int16 = 0
            For Each row As DataRow In dtTemp.Rows
                If row("STATUS").ToString = "NEW" Then
                    ttlnew = ttlnew + 1
                    If Convert.ToDouble(row("ITM_QTY").ToString) <= 0 Then
                        ttlnewnostk = ttlnewnostk + 1
                    End If
                End If
            Next
            lblNewCount.Text = ttlnew.ToString
            lblNewCountZeroStk.Text = ttlnewnostk.ToString
            gvCCItems.DataSource = dtTemp
            gvCCItems.DataBind()
        Else
            gvCCItems.DataSource = Nothing
            gvCCItems.DataBind()
            uiFun.displayMsg(Me, "", "Select Storer!", Session("gLang"))
        End If
    End Sub
    Protected Sub btnListCCItems_Click(sender As Object, e As EventArgs)
        getCCItems()
    End Sub


    Protected Sub btnExport_Click(sender As Object, e As EventArgs)
        Try
            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=CC_ITEMS_Export.xls")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            Using sw As New System.IO.StringWriter()
                Dim hw As New HtmlTextWriter(sw)
                gvCCItems.RenderControl(hw)
                Response.Output.Write(sw.ToString())
                Response.Flush()
                Response.[End]()
            End Using
        Catch ex As Exception
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        End Try
    End Sub
    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered
    End Sub
End Class
