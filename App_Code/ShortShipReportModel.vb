Imports Microsoft.VisualBasic

Public Class ShortShipReportModel
    Public Property DO_DATE() As String
    Public Property DO_Display_DATE() As String
    Public Property STO_NAME_CH() As String
    Public Property STO_ADDR1() As String
    Public Property ROUTE_ID() As String
    Public Property PLD_WH() As String
    Public Property DOD_UOM() As String
    Public Property PLD_LOC() As String
    Public Property PLD_SEQ() As String
    Public Property ITM_SKU_NO() As String
    Public Property PLD_ITEM_NO() As String
    Public Property PLD_EXPIRY_DATE() As String
    Public Property PLD_ITEM_QTY() As Double
    Public Property DOD_ITM_DESC() As String
    Public Property PLD_SS_QTY() As Double
    Public Property DOD_CUS_NAME() As String
    Public Property DOD_RULES() As String
    Public Property PLD_BATCH_NO() As String
    Public Property CUST_INFO() As String
    Public Property DOD_CO_CODE() As String

End Class

Public Class RootShortShipReportModel
    Public Property ShortShips() As List(Of ShortShipReportModel)
    Public Property ROUTE_ID() As String
    Public Property DO_DATE() As String
    Public Property DO_Display_DATE() As String
    Public Property STO_NAME_CH() As String
    Public Property STO_ADDR1() As String
End Class



