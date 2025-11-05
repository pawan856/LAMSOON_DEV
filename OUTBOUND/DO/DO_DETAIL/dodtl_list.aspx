<%@ Page Language="VB" AutoEventWireup="false" CodeFile="dodtl_list.aspx.vb" Inherits="OUTBOUND_DO_DO_DETAIL_dodtl_list" %>

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
        <div style="margin-top: -23px; margin-left: 116px;">
            <asp:Button ID="btnExportDataTable" runat="server" Text="Export Raw Data" />
        </div>

        <div>
            <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server"></asp:ToolkitScriptManager>
            <asp:Repeater ID="rptCustomers" runat="server">

                <ItemTemplate>
                    <div style="text-align: center; margin-left: 70px; margin-top: 31px; font-size: 25px; font-weight: bold">
                        <asp:Label ID="STO_ADDR1" runat="server" Text='<%# Eval("STO_ADDR1") %>' />
                        <div>货车出货日报表-总表</div>
                    </div>

                    <div style="padding-left: 260px; margin-top: 50px; display: flex; width: 750px; font-size: 25px; font-weight: bold">
                        <div class="col">
                            <asp:Label ID="lblDate" runat="server" Text="送货日期:" />
                            <asp:Label ID="DO_DATE" runat="server" Text='<%# Eval("DO_Display_DATE") %>' />
                            <br />
                            <asp:Label ID="lblRoute" runat="server" Text="车线:" />
                            <asp:Label ID="ROUTE_ID" runat="server" Text='<%# Eval("ROUTE_ID") %>' />

                        </div>

                        <div class="col" style="margin-left: auto;">

                            <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="width: 225px; text-align: center; font-size: 25px;">
                                <thead>
                                    <tr>
                                        <th colspan="2">貨車整潔</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>合格</td>
                                        <td>不合格</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <br />
                                        </td>
                                        <td></td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <br />
                    <center>

                          <table class="table table-bordered" border="1" cellpadding="0" cellspacing="0" bordercolor="#e1e1e0" style="font-family: Segoe UI; font-size: 25px; font-weight: bold; width: 1050PX; margin: -1px; border-collapse: collapse;">
                          <tr style="text-align: center;">
                                    <td style="width: 360px; border-collapse: collapse;">
                                        <asp:Label ID="lblSKU" runat="server" Text="产品编号" /></td>
                                   <%-- <td style="width: 220px">
                                        <asp:Label ID="lblITMDESC" runat="server" Text="产品描述" /></td>--%>
                                        <td style="width: 120px; border-collapse: collapse;"> 
                                        <asp:Label ID="lblLOTNo" runat="server" Text="批次" /></td>
                                     <%--<td style="width: 110px">  
                                        <asp:Label ID="lblLOCWH" runat="server" Text="位置" /></td>    --%>                                 
                                    <td style="width: 150px; border-collapse: collapse;">
                                        <asp:Label ID="lblEXPDate" runat="server" Text="到期日" /></td>   
                                    <td style="width: 180px; border-collapse: collapse;">  
                                        <asp:Label ID="lblRULES" runat="server" Text="最少保質期，   每次最多批次  不能超過上批次" /></td>
                                     <td style="width: 80px; border-collapse: collapse;">   
                                        <asp:Label ID="lblQTY" runat="server" Text="数量"/></td>
                                    <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="lblUOM" runat="server" Text="散装" /></td>
                                </tr>    
         
              <asp:Repeater ID="ddlInner" runat="server" DataSource='<%#Eval("Pickings") %>'>
                                <ItemTemplate> 
                                    
                                <tr style="text-align: center">
                                    <td style="width: 360px; border-collapse: collapse;">
                                        <asp:Label ID="ITM_SKU_NO" runat="server" Text='<%# Eval("ITM_SKU_NO") %>' /><br />
                                        <asp:Label ID="DOD_ITM_DESC" runat="server" Text='<%# Eval("DOD_ITM_DESC") %>' />
                                    </td>                   
                                    
                                    <td style="width: 120px; border-collapse: collapse;">   
                                        <asp:Label ID="DOD_BATCH_NO" runat="server" Text='<%# Eval("DOD_BATCH_NO") %>'  /><br />
                                         <asp:Label ID="Label1" runat="server" Text='<%# Eval("DOD_WH_CODE") + Eval("DOD_LOC_WH") %>' />
                                    </td>
                                     <%--<td style="width: 110px"> 
                                        <asp:Label ID="DOD_WH_CODE" runat="server" Text='<%# Eval("DOD_WH_CODE") + Eval("DOD_LOC_WH") %>' /></td>--%>
                                    
                                    <td style="width: 150px; border-collapse: collapse;"> 
                                        <asp:Label ID="DOD_EXPIRY_DATE" runat="server" Text='<%# Eval("DOD_EXPIRY_DATE") %>'  /></td>
                                   <td style="width: 180px; border-collapse: collapse;">
                                        <asp:Label ID="DOD_RULES" runat="server" Text='<%# Eval("DOD_RULES") %>'  /></td>
                                    <td style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="DOD_QTY" runat="server" Text='<%# Eval("DOD_QTY") %>' /></td>
                                    <td  style="width: 80px; border-collapse: collapse;">
                                        <asp:Label ID="DOD_UOM" runat="server" Text='<%# Eval("DOD_UOM") %>'  /></td>
                                </tr>  
                                    
                               </ItemTemplate>
                  </asp:Repeater>

                            <tr>
                            <td colspan="3"></td>                                    
                            <td>Total :</td>                                        
                            <td><asp:Label ID="txtTotalDODQty" runat="server" Text='<%#Eval("TOTAL_DOD_QTY") %>' /></td>
                            <td></td>
                            </tr>
                                  
            </table>            
                        
                         <div class="page-break"></div>

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
