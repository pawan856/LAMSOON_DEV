Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class MASTER_WM_del_dtl
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private db As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils("MAST_WM", Session("usr_id"), Me)
        ar.hideForm(Me)

        If Not IsPostBack Then
            ViewState("wh_code") = ""
            ViewState("wh_code") = Request("wh")

            ViewState("dt") = Nothing
            Call BindGV()

        End If
    End Sub

    Private Sub BindGV()
        Dim selectSQL As String = ""
        Dim wh_code As String = ""
        Dim dt As DataTable
        Dim paP As GlobalDBFunc.DBCmdPara

        If ViewState("wh_code") <> "" Then
            wh_code = ViewState("wh_code")
        Else
            wh_code = Request("wh")
        End If

        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = " select wh_code, fl_num, fl_name, '' as ar_code, '' as ar_name, '' as rk_code,'' as rk_name, '' as bn_code, 'F' as dtl_type from wms_wh_fl " & _
                    " Where wh_code=" & paP.AP(wh_code) & _
                    " union " & _
                    " select f1.wh_code,a1.fl_num,f1.fl_name,  a1.ar_code, a1.ar_name,'' as rk_code,'' as rk_name, '' as bn_code, 'A' as dtl_type from wms_wh_area a1, wms_wh_fl f1 " & _
                    " where a1.imp_code=f1.imp_code and a1.wh_code = f1.wh_code and a1.fl_num = f1.fl_num and a1.wh_code=" & paP.AP(wh_code) & _
                    " union " & _
                    " select f1.wh_code,rk1.fl_num,f1.fl_name,  rk1.ar_code, a1.ar_name,rk1.rk_code,rk1.rk_name, '' as bn_code, 'R' as dtl_type from wms_wh_area a1, wms_wh_fl f1, wms_wh_rack rk1 " & _
                    " where rk1.imp_code=f1.imp_code and rk1.wh_code = f1.wh_code and rk1.fl_num = f1.fl_num " & _
                         " AND rk1.imp_code=a1.imp_code and rk1.wh_code = a1.wh_code and rk1.fl_num = a1.fl_num and rk1.ar_code = a1.ar_code and rk1.wh_code=" & paP.AP(wh_code) & _
                    " union " & _
                    " select f1.wh_code,b1.fl_num,f1.fl_name,  b1.ar_code, a1.ar_name,b1.rk_code,rk1.rk_name, b1.bn_code, 'B' as dtl_type from wms_wh_area a1, wms_wh_fl f1, wms_wh_rack rk1, wms_wh_bin b1 " & _
                    " where b1.imp_code=f1.imp_code and b1.wh_code = f1.wh_code and b1.fl_num = f1.fl_num " & _
                         " AND b1.imp_code=a1.imp_code and b1.wh_code = a1.wh_code and b1.fl_num = a1.fl_num and b1.ar_code = a1.ar_code " & _
                         " AND b1.imp_code=rk1.imp_code and b1.wh_code = rk1.wh_code and b1.fl_num = rk1.fl_num and b1.ar_code = rk1.ar_code AND b1.rk_code = rk1.rk_code and b1.wh_code=" & paP.AP(wh_code) & _
                    " order by wh_code,fl_num, ar_code, rk_code, bn_code "

        dt = gDB.getDataTable(selectSQL, , , , paP)

        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case e.CommandName
            Case "DEL"
                Dim wh_code, fl_num, ar_code, rk_code, bn_code, dtl_type As String
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                Dim deleteSQL As String = ""

                wh_code = ViewState("wh_code")
                fl_num = DirectCast(gvRow.FindControl("fl_num"), HiddenField).Value
                ar_code = DirectCast(gvRow.FindControl("ar_code"), HiddenField).Value
                rk_code = DirectCast(gvRow.FindControl("rk_code"), HiddenField).Value
                bn_code = DirectCast(gvRow.FindControl("bn_code"), HiddenField).Value
                dtl_type = DirectCast(gvRow.FindControl("dtl_type"), HiddenField).Value

                Dim paraWH(4) As String

                paraWH(0) = dtl_type
                paraWH(1) = fl_num
                paraWH(2) = ar_code
                paraWH(3) = rk_code
                paraWH(4) = bn_code

                Dim gConn As SqlConnection
                Dim paP As GlobalDBFunc.DBCmdPara

                If ValidateAll(paraWH) Then
                    gConn = gDB.getConnection()

                    Dim transaction As SqlTransaction
                    transaction = gConn.BeginTransaction()

                    Try
                        Select Case dtl_type
                            Case "F"
                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_fl where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_area where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) 

                                gDB.amendData(deleteSQL, gConn, transaction, paP)

                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_rack where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) 

                                gDB.amendData(deleteSQL, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_bin where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)
                            Case "A"
                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_area where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)

                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_rack where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_bin where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)

                            Case "R"
                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_rack where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code) & " AND rk_code=" & paP.AP(rk_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)


                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_bin where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code) & " AND rk_code=" & paP.AP(rk_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)

                            Case "B"
                                paP = New GlobalDBFunc.DBCmdPara
                                deleteSQL = "Delete from wms_wh_bin where imp_code=" & paP.AP(Session("imp_code")) & " AND wh_code=" & paP.AP(wh_code) & " AND fl_num=" & paP.AP(fl_num) & _
                                            " AND ar_code=" & paP.AP(ar_code) & " AND rk_code=" & paP.AP(rk_code) & " AND bn_code=" & paP.AP(bn_code)

                                gDB.amendData(deleteSQL, gConn, transaction, paP)

                        End Select

                        transaction.Commit()
                        Dim alert_msg As String = ""

                        Select Case dtl_type
                            Case "F"
                                alert_msg = "Selected Floor and Area/Rack/Bin of this floor has been deleted."
                            Case "A"
                                alert_msg = "Selected Area and Rack/Bin in this area has been deleted."
                            Case "R"
                                alert_msg = "Selected Rack and Bin in this rack has been deleted."
                            Case "B"
                                alert_msg = "Selected Bin has been deleted."
                        End Select

                        uiFun.displayMsg(Me, "", alert_msg, Session("gLang"))
                    Catch ex As Exception
                        transaction.Rollback()
                        uiFun.displayMsg(Me, "1002", "", Session("gLang"))

                    Finally
                        If gConn IsNot Nothing Then
                            If gConn.State = ConnectionState.Open Then
                                gConn.Close()
                                gConn.Dispose()
                            End If
                        End If
                    End Try

                End If

                Call BindGV()

        End Select
    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
         Select e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("fl_num"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "fl_num").ToString.Trim
                CType(e.Row.FindControl("ar_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ar_code").ToString.Trim
                CType(e.Row.FindControl("rk_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rk_code").ToString.Trim
                CType(e.Row.FindControl("bn_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "bn_code").ToString.Trim
                CType(e.Row.FindControl("dtl_type"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "dtl_type").ToString.Trim

                Dim alertMsg As String = ""
                Select Case DataBinder.Eval(e.Row.DataItem, "dtl_type").ToString.Trim
                    Case "F"
                        alertMsg = "This Floor and all Area/Rack/Bin of this Floor will be deleted.\n Confirm Delete? "
                    Case "A"
                        alertMsg = "This Area and all Rack/Bin of this Area will be deleted.\n Confirm Delete? "
                    Case "R"
                        alertMsg = "This Rack and all Bin of this Rack will be deleted.\n Confirm Delete? "
                    Case "B"
                        alertMsg = "This Bin will be deleted.\n Confirm Delete? "
                End Select

                CType(e.Row.FindControl("btnDelDtl"), Button).OnClientClick = "if (!confirm('" & alertMsg & "')) return false;"
        End Select
    End Sub

    Private Function ValidateAll(ByVal para() As String) As Boolean
        Dim wh_code, fl_num, ar_code, rk_code, bn_code, dtl_type As String
        wh_code = ViewState("wh_code")

        dtl_type = para(0)
        fl_num = para(1)
        ar_code = para(2)
        rk_code = para(3)
        bn_code = para(4)

        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim tempCount As Integer = 0

        paP = New GlobalDBFunc.DBCmdPara
        Select Case dtl_type
            Case "F"
                selectSQL = "Select Count(*) from wms_item_loc_bal where imp_code=" & paP.AP(Session("imp_code")) & " AND ILOC_WH=" & paP.AP(wh_code) & " AND ILOC_FLOOR=" & paP.AP(fl_num)  & " AND ILOC_BAL_QTY > 0 "

            Case "A"
                selectSQL = "Select Count(*) from wms_item_loc_bal where imp_code=" & paP.AP(Session("imp_code")) & " AND ILOC_WH=" & paP.AP(wh_code) & " AND ILOC_FLOOR=" & paP.AP(fl_num) & _
                            " AND ILOC_AREA=" & paP.AP(ar_code) & " AND ILOC_BAL_QTY > 0 "

            Case "R"
                selectSQL = "Select Count(*) from wms_item_loc_bal where imp_code=" & paP.AP(Session("imp_code")) & " AND ILOC_WH=" & paP.AP(wh_code) & " AND ILOC_FLOOR=" & paP.AP(fl_num) & _
                            " AND ILOC_AREA=" & paP.AP(ar_code) & " AND ILOC_RACK=" & paP.AP(rk_code) & " AND ILOC_BAL_QTY > 0 "

            Case "B"
                selectSQL = "Select Count(*) from wms_item_loc_bal where imp_code=" & paP.AP(Session("imp_code")) & " AND ILOC_WH=" & paP.AP(wh_code) & " AND ILOC_FLOOR=" & paP.AP(fl_num) & _
                            " AND ILOC_AREA=" & paP.AP(ar_code) & " AND ILOC_RACK=" & paP.AP(rk_code) & " AND ILOC_BIN=" & paP.AP(bn_code) & " AND ILOC_BAL_QTY > 0 "
            Case Else
                selectSQL = ""
        End Select

        If selectSQL <> "" Then
            tempCount = gU.decodeEmptyCInt(db.getValueFromSQL(selectSQL, , , paP), 0)
        End If

        If tempCount > 0 Then
            uiFun.displayMsg(Me, "", "The selected item cannot be deleted!\nLocation Balance exists in this location.", Session("gLang"))
            Return False
        End If


        Return True
    End Function
  
End Class
