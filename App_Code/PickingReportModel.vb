Imports Microsoft.VisualBasic

Public Class PickingReportModel
    Public Property DO_DATE() As DateTime
    Public Property DO_Display_DATE() As String
    Public Property STO_NAME_CH() As String
    Public Property STO_ADDR1() As String
    Public Property ROUTE_ID() As String
    Public Property DOD_WH_CODE() As String
    Public Property DOD_BATCH_NO() As String
    Public Property DOD_UOM() As String
    Public Property DOD_LOC_WH() As String
    Public Property DOD_SEQ() As String
    Public Property ITM_SKU_NO() As String
    Public Property DOD_EXPIRY_DATE() As String
    Public Property DOD_QTY() As String
    Public Property DOD_ITM_DESC() As String
    Public Property DOD_SS_QTY() As Double
    Public Property DOD_CUS_NAME() As String
    Public Property DOD_RULES() As String
    Public Property DOD_REM() As String
    Public Property DOD_FL_CODE() As String
    Public Property TOTAL_DOD_QTY() As String

End Class


Public Class RootPickingReportModel
    Public Property Pickings() As List(Of PickingReportModel)
    Public Property ROUTE_ID() As String
    Public Property DO_Display_DATE() As String
    Public Property DO_DATE() As String
    Public Property STO_NAME_CH() As String
    Public Property STO_ADDR1() As String
    Public Property TOTAL_DOD_QTY() As String
End Class

