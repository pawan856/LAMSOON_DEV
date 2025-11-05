Imports Microsoft.VisualBasic

Public Class ITMBALRSVDReportModel
    Public Property WH_CODE() As String
    Public Property wh_loc() As String
    Public Property lot_no() As String
    Public Property ITM_CODE() As String
    Public Property ITM_DESC() As String
    Public Property ITM_SKU_NO() As String
    Public Property ITM_SHELF_LIFE() As String
    Public Property Stock_Balance() As String
    Public Property Allocated() As String
    Public Property Balance() As String
    Public Property Total_Stock_Balance() As String
    Public Property Total_Allocated() As String
    Public Property Total_Balance() As String
    Public Property ILOC_EXPIRY_DATE() As String
    Public Property ITM_UOM() As String

End Class

Public Class RootITMBALRSVDReportModel
    Public Property ITMBALRSVD() As List(Of ITMBALRSVDReportModel)
    Public Property ITM_SKU_NO() As String
    Public Property ITM_DESC() As String
    Public Property Total_Stock_Balance() As String
    Public Property Total_Allocated() As String
    Public Property Total_Balance() As String
End Class







