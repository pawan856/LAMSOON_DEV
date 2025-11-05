<%@ Page Language="VB" AutoEventWireup="false" CodeFile="short_ship.aspx.vb" Inherits="OUTBOUND_DO_SHORT_SHIP_short_ship" %>

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

    <script type="text/javascript">  

        function printExternal() {
            document.getElementById("btnExport").style.display = "none";
            document.getElementById("btnExportDataTable").style.display = "none";
            window.print();  
         }   

        function printAfter() {          
            document.getElementById("btnExport").style.display = "block";
            document.getElementById("btnExportDataTable").style.display = "block";
        }
       
    </script>

</head>

<body onload="printExternal();" onafterprint="printAfter();">
    <form id="form1" runat="server">
        <asp:Button ID="btnExport" runat="server" Text="Export To Excel" />
        <div style="margin-top: -23px;margin-left: 116px;">
              <asp:Button ID="btnExportDataTable" runat="server" Text="Export Raw Data" />
        </div>
       
              
       
        <div>
            <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></asp:ToolkitScriptManager>
            <asp:Repeater ID="rptCustomers" runat="server">
                
                <ItemTemplate>                    
                    <div style="text-align:center;margin-left: 70px;margin-top: 8px;font-size: 25px;font-weight:bold">
                           <asp:Label ID="STO_ADDR1" runat="server" Text='<%# Eval("STO_ADDR1") %>'/>
                             <div>缺貨日報表</div>
                       </div>                        
                    
                    <div style="padding-left: 260px;margin-top: 45px;font-size: 25px; font-weight:bold">
                            <asp:Label ID="lblDate" runat="server" Text="送货日期:" />
                            <asp:Label ID="DO_DATE" runat="server" Text='<%# Eval("DO_Display_DATE") %>' />  
                      <br />    
                            <asp:Label ID="lblRoute" runat="server" Text="车线:"/> 
                            <asp:Label ID="ROUTE_ID" runat="server" Text='<%# Eval("ROUTE_ID") %>'/> 
                     </div>                             
                
                   <br />
                    <center>
                   
                    <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 25px; font-weight: bold; width: 1050PX; margin: -1px; border-collapse: collapse;">
                          <tr style="text-align: center">
                                    <td style="width: 250px; border-collapse: collapse;">
                                        <asp:Label ID="lblSKU" runat="server" Text="产品编号" /></td>
                                    <%--<td style="width: 240px; border-collapse: collapse;">
                                        <asp:Label ID="lblITMDESC" runat="server" Text="e产品描述" /></td>--%>
                                    <td style="width: 110px; border-collapse: collapse;">
                                        <asp:Label ID="lblLOCWH" runat="server" Text="位置" /></td>
                               <td style="width: 110px; border-collapse: collapse;">
                                        <asp:Label ID="lblLOTNO" runat="server" Text="批次" /></td>
                                    <%--<td style="width: 300px">
                                        <asp:Label ID="lblRULES" runat="server" Text="包装说明" /></td>--%>
                                <td style="width: 300px; border-collapse: collapse;">
                                        <asp:Label ID="lblCustomer" runat="server" Text="客户订单资料" /></td>
                                    <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="lblQTY" runat="server" Text="数量"/></td>
                                    <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="lblUOM" runat="server" Text="散装" /></td>
                               <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="lblSS" runat="server" Text="缺貨數量" /></td>
                                </tr>
                     
    
                        <asp:Repeater ID="ddlInner" runat="server" DataSource='<%#Eval("ShortShips") %>'>
                         <ItemTemplate>   
                                     <tr style="text-align: center">
                                    <td style="width: 270px; border-collapse: collapse;">
                                        <asp:Label ID="ITM_SKU_NO" runat="server" Text='<%# Eval("ITM_SKU_NO") %>' /><br />
                                        <asp:Label ID="DOD_ITM_DESC" runat="server" Text='<%# Eval("DOD_ITM_DESC") %>' />
                                    </td>
                                    <%--<td style="width: 240px; border-collapse: collapse;">
                                        </td>--%>
                                    <td style="width: 100px; border-collapse: collapse;">
                                        <asp:Label ID="PLD_LOC" runat="server" Text='<%# Eval("PLD_WH") + Eval("PLD_LOC") %>' /></td>
                                  <td style="width: 100px; border-collapse: collapse;">
                                        <asp:Label ID="PLD_BATCH_NO" runat="server" Text='<%# Eval("PLD_BATCH_NO") %>' /></td>
                                    <%--  <td style="width: 300px">
                                        <asp:Label ID="DOD_RULES" runat="server" Text='<%# Eval("DOD_RULES") %>'  /></td>--%>
                                    <td style="width: 300px; border-collapse: collapse;">
                                     <asp:Label ID="CUST_INFO" runat="server" Text='<%# Eval("CUST_INFO") %>' /></td>
                                    <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="PLD_ITEM_QTY" runat="server" Text='<%# Eval("PLD_ITEM_QTY") %>' /></td>
                                    <td  style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="DOD_UOM" runat="server" Text='<%# Eval("DOD_UOM") %>'  /></td>
                                    <td  style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="PLD_SS_QTY" runat="server" Text='<%# Eval("PLD_SS_QTY") %>'  /></td>
                                </tr>                         
                             
                        </ItemTemplate>
              </asp:Repeater>
                        
                    

                             </table>

                   <div class="page-break">

                         </div>     

                  </center>

                </ItemTemplate>

            </asp:Repeater>

        </div>

        <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>

    </form>
</body>
</html>


