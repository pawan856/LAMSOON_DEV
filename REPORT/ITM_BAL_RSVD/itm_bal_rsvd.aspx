<%@ Page Language="VB" AutoEventWireup="false" CodeFile="itm_bal_rsvd.aspx.vb" Inherits="REPORT_ITM_BAL_RSVD_itm_bal_rsvd" %>

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
    </style>

</head>

<body onload="window.print()">

    <form id="form1" runat="server">
        <asp:Button ID="btnExport" runat="server" Text="Export to excel" />&nbsp;&nbsp;&nbsp;
         <asp:Button ID="btnExportDataTable" runat="server" Text="Export Raw Data" />

        <div>
            <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></asp:ToolkitScriptManager>
            <asp:Repeater ID="rptCustomers" runat="server">
                <HeaderTemplate>
                    <center>
                     <div style="font-weight:bold;font-size:18px">       
                            <asp:Label ID="lblHeaderName" runat="server" Text="Item Balance Report" />                            
                       </div> 
                  
                     <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 1100px; margin: -1px">
                         
                         <tr style="text-align: center">
                            <td style="width:132px"><asp:Label ID="lblSKU" runat="server" Text="货品编号" /></td>
                            <td  style="width:83px"><asp:Label ID="lblITMDESC" runat="server" Text="货品名称" /></td>  
                                 <td colspan="7"></td>                                
                           </tr>
                </HeaderTemplate>

                <ItemTemplate>
                    <br />

                    <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 1100px; margin: -1px">
                        <tr style="text-align: center">
                            <td style="width: 260px">
                                <asp:Label ID="ITM_SKU_NO" runat="server" Text='<%# Eval("ITM_SKU_NO") %>' /></td>
                            <td style="width: 200px">
                                <asp:Label ID="ITM_DESC" runat="server" Text='<%# Eval("ITM_DESC") %>' /></td>
                            <td>

                                <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 12px; font-weight: bold; width: 880px; margin: -1px">
                                    <tr style="text-align: center">
                                        <td style="width: 102px">
                                            <asp:Label ID="lblWH" runat="server" Text="子库存" /></td>
                                        <td style="width: 107px">
                                            <asp:Label ID="lblLOT" runat="server" Text="Lot Number" /></td>
                                        <td style="width: 103px">
                                            <asp:Label ID="lblMSL" runat="server" Text="產品保質天數" /></td>
                                        <td style="width: 100px">
                                            <asp:Label ID="lblSTKBAL" runat="server" Text="Stock Balance" /></td>
                                        <td style="width: 100px">
                                            <asp:Label ID="lblAllocated" runat="server" Text="分配數量" /></td>
                                        <td style="width: 100px">
                                            <asp:Label ID="lblBAL" runat="server" Text="分配餘數" /></td>
                                        <td style="width: 103px">
                                            <asp:Label ID="lblExpDate" runat="server" Text="產品到期日" /></td>
                                        <td style="width: 103px">
                                            <asp:Label ID="lblUOM" runat="server" Text="UOM" /></td>

                                        <%--<td style="width: 105px">
                                            <asp:Label ID="lblLOC" runat="server" Text="Location" /></td>--%>
                                    </tr>

                                    <asp:Repeater ID="ddlInner" runat="server" DataSource='<%#Eval("ITMBALRSVD") %>'>
                                        <ItemTemplate>
                                            <tr>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="WH_CODE" runat="server" Text='<%# Eval("WH_CODE") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="lot_no" runat="server" Text='<%# Eval("lot_no") %>' /></td>
                                                <td style="width: 150px;border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="ITM_SHELF_LIFE" runat="server" Text='<%# Eval("ITM_SHELF_LIFE") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="Stock_Balance" runat="server" Text='<%# Eval("Stock_Balance") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="Allocated" runat="server" Text='<%# Eval("Allocated") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="Balance" runat="server" Text='<%# Eval("Balance") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="ILOC_EXPIRY_DATE" runat="server" Text='<%# Eval("ILOC_EXPIRY_DATE") %>' /></td>
                                                <td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="ITM_UOM" runat="server" Text='<%# Eval("ITM_UOM") %>' /></td>


                                                <%--<td style="width: 150px; border: 1px solid #e1e1e0; border-collapse: collapse;">
                                                    <asp:Label ID="wh_loc" runat="server" Text='<%# Eval("wh_loc") %>' /></td>--%>
                                            </tr>
                                        </ItemTemplate>

                                    </asp:Repeater>

                                    <tr>
                                        <td colspan="2"></td>
                                        <td>Total : </td>
                                        <td>
                                            <asp:Label ID="txtTotalStockBalance" runat="server" Text='<%#Eval("Total_Stock_Balance") %>' />
                                        </td>
                                        <td>
                                            <asp:Label ID="txtTotalAllocated" runat="server" Text='<%#Eval("Total_Allocated") %>' />
                                        </td>
                                        <td>
                                            <asp:Label ID="txtTotalBalance" runat="server" Text='<%#Eval("Total_Balance") %>' />
                                        </td>
                                    </tr>

                                    <tr>
                                        <td colspan="5">
                                            <asp:Label ID="lblStockBalance" runat="server" />

                                        </td>
                                    </tr>

                                </table>
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
