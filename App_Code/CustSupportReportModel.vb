Imports Microsoft.VisualBasic

Public Class CustSupportReportModel
    Public Property SKU_NUMBER() As String
    Public Property SKU_NAME() As String
    Public Property CUS_NAME() As String
    Public Property CO_SENDER_REGION() As String
    Public Property DOD_MIN_SHELF_LIFE() As String
    Public Property DOD_MAX_LOT() As String
    Public Property DOD_MIN_PROD_DATE() As String
    Public Property COD_WH_CODE() As String
    Public Property COD_PALLET_NO() As String
    Public Property COD_QTY() As String
    Public Property ALLOCATED() As String
    Public Property DOD_CO_CODE() As String
    Public Property CUS_ORDER() As String
    Public Property CO_INV_NO() As String
    Public Property Total_QTY() As String

    Public Property IO_CODE As String
    Public Property CUST_PO_NUMBER As String
    Public Property COD_Plant As String
    Public Property CUS_CODE As String
    Public Property COD_Carton_no As String
    Public Property Formula As String
    Public Property COD_UOM As String
    Public Property CO_DATE As String
    Public Property ROUTE_ID As String
End Class

Public Class RootCustSupportReportModel
    Public Property CustSupports() As List(Of CustSupportReportModel)
    Public Property SKU_NUMBER() As String
    Public Property SKU_NAME() As String
    Public Property Total_QTY() As String
End Class

Public Class DOD_CO_CODEModel
    Public Property DOD_CO_CODE() As String

End Class






