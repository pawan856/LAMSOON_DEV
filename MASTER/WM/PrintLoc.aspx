<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PrintLoc.aspx.vb" Inherits="MASTER_WM_PrintLoc" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Export Location Code</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">
    


</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table  border="0" cellspacing="1" cellpadding="1" align="center" style="width: 100%">
             <tr>
                    <td class="TITLE">
                        <b><asp:Label ID="lheader" runat="server" text="Export Location Code to Excel" /></b>
                    </td>
                </tr>
                <tr>
                    <td class="menuTD">
                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="menuTD" width="50%">
                                <asp:Button runat="server" id="btnPrint1" CssClass="all_button" Text="Export" />
                                <asp:Button runat="server" id="btnPrint3" CssClass="all_button" Text="Export without Barcode" />
                                <input type="button" value="Close" onclick="javascript:window.close();" class="all_button" />
                            </td>                            
                        </tr>
                        </table>
                    </td>
               </tr>
               <tr>
                    <td>
                        <asp:Button runat="server" ID="btnSel" text="Select All" CssClass="all_button"/>
                        <asp:Button runat="server" ID="btnUnSel" text="Un-Select All" CssClass="all_button"/>
                    </td>
               </tr>
               <tr>
                  <td>
                      <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                             Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False"  EmptyDataText="No Bin is Found in This Warehouse!"
                             BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                             CellPadding="3" CaptionAlign="Top" DataKeyNames="BN_CODE" HorizontalAlign="Left">
                             <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                             <RowStyle CssClass="GV" />
                             <Columns>
                                <asp:TemplateField ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="3%">
                                   <ItemTemplate>
                                       <asp:CheckBox runat="server" id="nSelect" />
                                       <asp:HiddenField runat="server" ID="wh_code" />
                                       <asp:HiddenField runat="server" ID="ar_code" />
                                       <asp:HiddenField runat="server" ID="fl_num"  /> 
                                       <asp:HiddenField runat="server" ID="rk_code" />                                      
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Warehouse" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="wh_name" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Floor" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="FL_NAME" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Area" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="AR_NAME" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Rack" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="RK_NAME" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Bin" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="BN_CODE" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="EDI Code" ItemStyle-Font-Size="11px" HeaderStyle-HorizontalAlign="Left" >
                                   <ItemTemplate>
                                     <asp:Label runat="server" ID="BN_CSMS_CODE" /> 
                                   </ItemTemplate>                                                                           
                                </asp:TemplateField>
                         </Columns>
                      </asp:GridView>
                  </td>
               </tr>    
               <tr>
                  <td class="menuTD">
                     <table border="0" cellspacing="0" cellpadding="0" width="100%">
                      <tr>
                         <td class="menuTD" width="50%">
                               <asp:Button runat="server" id="btnPrint2" CssClass="all_button" Text="Export" />
                               <input type="button" value="Close" onclick="javascript:window.close();" class="all_button" />
                           </td>                            
                       </tr>
                      </table>
                    </td>
               </tr>    
        </table>
    </div>
    </form>
</body>
</html>
