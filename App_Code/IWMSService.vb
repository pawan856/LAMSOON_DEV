Imports System.ServiceModel
Imports System
Imports System.Configuration
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports System.Data

' NOTE: If you change the class name "IWMSService" here, you must also update the reference to "IWMSService" in Web.config.
<ServiceContract()> _
Public Interface IWMSService

    <OperationContract()> _
    Function LoginWMS(ByVal user As String, ByVal pwd As String) As Boolean

    <OperationContract()> _
    Function getROList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function getROItemList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function getGRList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function getGRItemList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function getPickList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function getPickItemList(ByVal storer As String) As DataTable

    <OperationContract()> _
    Function updatePickList(ByVal usrID As String, ByVal dt As DataTable) As Boolean

    <OperationContract()> _
    Function getLocation() As DataTable

    <OperationContract()> _
    Function getDataTableFrmService(ByVal as_sql As String) As DataTable

    <OperationContract()> _
    Function getTableSchema(ByVal as_sql As String) As DataTable

    <OperationContract()> _
    Function getValueFromSQL(ByVal as_sql As String) As String

    <OperationContract()> _
    Function createGR(ByVal grTable As DataTable, ByVal roTable As DataTable, _
                      ByVal usrID As String, ByVal gr_date As String, _
                      ByVal paTable As DataTable) As Boolean
End Interface
