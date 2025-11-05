<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cust_support.aspx.vb" Inherits="REPORT_CUST_SUPPORT_cust_support" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <style>
        @media all {
            .page-break {
                display: none;
            }
        }

        @media print {
            .page-break {
                display: block;
                page-break-before: always;
            }
        }

        table, th, td {
            border: 1px solid black;
            border-collapse: collapse;
        }
    </style>

</head>

<body onload="window.print()">
    <form id="form1" runat="server">
        <asp:Button ID="btnExport" runat="server" Text="Export to excel" />
        &nbsp;&nbsp;&nbsp;
         <asp:Button ID="btnExportDataTable" runat="server" Text="Export Raw Data" />

        <div>
            <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></asp:ToolkitScriptManager>
            <asp:Repeater ID="rptCustomers" runat="server">
                <HeaderTemplate>
                    <center>
                     <div style="font-weight:bold;font-size:18px">       
                            <asp:Label ID="lblHeaderName" runat="server" Text="Sales Admin Shortage Report" />                            
                       </div> 

                    <table class="table table-bordered"  cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 1100px; margin: -1px;border: 1px solid black;border-collapse: collapse;">
                          <tr style="text-align: center">
                           <td style="width: 140px;border: 1px solid black; border-collapse: collapse;">
                                <asp:Label ID="lblSKU" runat="server" Text="貨品編號" /></td>
                            <td style="width: 74px;border: 1px solid black; border-collapse: collapse;">
                                <asp:Label ID="lblITMDESC" runat="server" Text="貨品名稱" /></td> 
                                 <td colspan="10"></td>                                
                           </tr>
                </HeaderTemplate>

                <ItemTemplate>
                    <br />

                    <table class="table table-bordered" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 1100px; margin: -1px; border: 1px solid black; border-collapse: collapse;">
                        <tr style="text-align: center;">
                            <td style="width: 260px; border: 1px solid black; border-collapse: collapse;">
                                <asp:Label ID="SKU_NUMBER" runat="server" Text='<%# Eval("SKU_NUMBER") %>' /></td>
                            <td style="width: 200px; border: 1px solid black; border-collapse: collapse;">
                                <asp:Label ID="SKU_NAME" runat="server" Text='<%# Eval("SKU_NAME") %>' /></td>
                            <td>

                                <table class="table table-bordered" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 880px; margin: -1px; border: 1px solid black; border-collapse: collapse;">
                                    <tr style="text-align: center">
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="IO_CODE" runat="server" Text="倉庫編號" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="Formula" runat="server" Text="Release Check" /></td>
                                        <td style="width: 100px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblWH" runat="server" Text="子庫存" /></td>
                                         <td style="width: 200px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblCUSORDER" runat="server" Text="訂單編號" /></td>
                                         <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="COD_Plant" runat="server" Text="訂單行號" /></td>
                                         <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="CUST_PO_NUMBER" runat="server" Text="客戶訂單號" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="CUS_CODE" runat="server" Text="客戶編號" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblCUSNAME" runat="server" Text="客戶名稱" /></td>  
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="COD_Carton_no" runat="server" Text="銷售員" /></td>
                                        <td style="width: 191px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblCRHLD" runat="server" Text="Credit Hold" /></td>
                                        <td style="width: 100px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblAllocated" runat="server" Text="Allocated Fail" /></td>
                                         <td style="width: 100px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblQTY" runat="server" Text="訂購數量" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="COD_UOM" runat="server" Text="單位" /></td>
                                        <td style="width: 200px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblMPD" runat="server" Text="客戶後熟天數要求" /></td>
                                        <td style="width: 200px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblMSL" runat="server" Text="客戶最少保質天數要求" /></td>
                                         <td style="width: 110px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblLASTLOT" runat="server" Text="客戶最後接受批次" /></td>
                                         <td style="width: 190px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="lblMAXLOT" runat="server" Text="客戶可接受最多批次" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="CO_DATE" runat="server" Text="計劃發貨日期" /></td>
                                        <td style="width: 170px; border: 1px solid black; border-collapse: collapse;">
                                            <asp:Label ID="ROUTE_ID" runat="server" Text="車線" /></td>                                                                              
                                    </tr>

                                    <asp:Repeater ID="ddlInner" runat="server" DataSource='<%#Eval("CustSupports") %>'>
                                        <ItemTemplate>
                                            <tr class="table table-bordered" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 879px; margin: -1px; border-collapse: collapse;">
                                                <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="IO_CODE" runat="server" Text='<%# Eval("IO_CODE") %>' /></td>
                                                <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="Formula" runat="server" Text='<%# Eval("Formula") %>' /></td>
                                                 <td style="width: 100px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_WH_CODE" runat="server" Text='<%# Eval("COD_WH_CODE") %>' /></td>
                                                 <td style="width: 139px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CUS_ORDER" runat="server" Text='<%# Eval("CO_INV_NO") %>' /></td>
                                                 <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_Plant" runat="server" Text='<%# Eval("COD_Plant") %>' /></td>
                                                <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CUST_PO_NUMBER" runat="server" Text='<%# Eval("CUST_PO_NUMBER") %>' /></td> 
                                                 <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CUS_CODE" runat="server" Text='<%# Eval("CUS_CODE") %>' /></td>
                                                <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CUS_NAME" runat="server" Text='<%# Eval("CUS_NAME") %>' /></td>  
                                                <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_Carton_no" runat="server" Text='<%# Eval("COD_Carton_no") %>' /></td>
                                                 <td style="width: 132px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CO_SENDER_REGION" runat="server" Text='<%# Eval("CO_SENDER_REGION") %>' /> </td>
                                                 <td style="width: 96px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="ALLOCATED" runat="server" Text='<%# Eval("ALLOCATED") %>' />
                                                     <td style="width: 59px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_QTY" runat="server" Text='<%# Eval("COD_QTY") %>' />
                                                         <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_UOM" runat="server" Text='<%# Eval("COD_UOM") %>' /></td>
                                                         <td style="width: 128px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="DOD_MIN_PROD_DATE" runat="server" Text='<%# Eval("DOD_MIN_PROD_DATE") %>' /></td>
                                                          <td style="width: 133px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="DOD_MIN_SHELF_LIFE" runat="server" Text='<%# Eval("DOD_MIN_SHELF_LIFE") %>' /></td>
                                                         <td style="width: 79px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="COD_PALLET_NO" runat="server" Text='<%# Eval("COD_PALLET_NO") %>' /></td>
                                                          <td style="width: 143px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="DOD_MAX_LOT" runat="server" Text='<%# Eval("DOD_MAX_LOT") %>' /></td>
                                                          <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="CO_DATE" runat="server" Text='<%# Eval("CO_DATE") %>' /></td>
                                                         <td style="width: 130px; border: 1px solid black; border-collapse: collapse;">
                                                    <asp:Label ID="ROUTE_ID" runat="server" Text='<%# Eval("ROUTE_ID") %>' /></td>                                                 
                                                </td>
                                               
                                                </td>
                                            </tr>

                                        </ItemTemplate>
                                    </asp:Repeater>

                                    <tr>
                                        <td colspan="10"></td>
                                        <td>Total Qty:  </td>
                                        <td>
                                            <asp:Label ID="txtTotal" runat="server" Text='<%#Eval("Total_QTY") %>' />
                                        </td>
                                    </tr>

                                </table>
                            </td>
                        </tr>
                    </table>

                </ItemTemplate>

                <FooterTemplate>
                    </table>
                    </center>
                </FooterTemplate>

            </asp:Repeater>

        </div>
        <asp:GridView ID="GridTableData" runat="server" AutoGenerateColumns="True" Visible="false"></asp:GridView>
        <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>

    </form>

</body>

</html>
