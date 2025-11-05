<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ROLOOKUP.aspx.vb" Inherits="OUTBOUND_CO_ROLOOKUP" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>RO Lookup</title>
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
    <form id="myform" runat="server">
    <input type="hidden" name="moduleAction" value=""/>
    <div>   
    <br />
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
             <tr>
                <td colspan="4" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                    <asp:label ID="lheader" runat ="server" Text="RO Lookup" />
                 </td>
             </tr> 
             <tr>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_storer" Text="Storer" />:
                    </font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label runat="server" ID="storer_name" />
                        <asp:HiddenField runat="server" ID="storer_code" />
                        <asp:HiddenField runat="server" ID="co_code" />
                        <asp:hiddenfield runat="server" ID="cod_seq" />
                    </font>
                </td>
             </tr>
             <tr>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_item_code" Text="Item Code" />:
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label runat="server" ID="item_code" />
                    </font>
                </td>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_batch_no" Text="Batch No" />:
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label runat="server" ID="batch_no" />
                    </font>
                </td>
             </tr>
             <tr>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_pack_key" Text="Pack Key" />:
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label runat="server" ID="pack_key" />
                    </font>
                </td>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_pallet_no" Text="Pallet No" />:
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label runat="server" ID="pallet_no" />
                    </font>
                </td>
             </tr>
        </table>
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
         Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False"  EmptyDataText="No RO is found for the specific item."
         BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
         CellPadding="3" CaptionAlign="Top" DataKeyNames="ROD_SEQ" HorizontalAlign="Left">
         <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
         <Columns>
             <asp:TemplateField HeaderText="RO Code"  ItemStyle-Font-Size="11px">
                <ItemTemplate>
                    <asp:Label runat="server" ID="ro_code" />
                </ItemTemplate>
             </asp:TemplateField> 
             <asp:TemplateField HeaderText="RO Seq."  ItemStyle-Font-Size="11px">
                <ItemTemplate>
                    <asp:Label runat="server" ID="rod_seq" />
                </ItemTemplate>
             </asp:TemplateField>
             <asp:TemplateField HeaderText="RO Qty."  ItemStyle-Font-Size="11px">
                <ItemTemplate>
                    <asp:Label runat="server" ID="rod_qty" />
                </ItemTemplate>
             </asp:TemplateField>
             <asp:BoundField HeaderText="Manu. Date" DataField="rod_manu_date" ItemStyle-Font-Size="11px"  />   
             <asp:BoundField HeaderText="Expirty Date" DataField="rod_expiry_date" ItemStyle-Font-Size="11px"  />  
             <asp:TemplateField >
                <ItemTemplate>
                    <asp:Button runat="server" ID="btnSel" CommandName="SELECT" text="Select" CssClass="all_button" />
                </ItemTemplate>
             </asp:TemplateField> 
         </Columns>
        </asp:GridView>
    </div>
    </form>
</body>
</html>
