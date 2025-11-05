<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Availability.aspx.vb" Inherits="Availability" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Availability</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">
function DisableDeleteButton() {
    var grp = document.getElementsByName("btnDelete");
    var count;
    count = grp.length;
    if (count == 1) {
       grp[0].disabled = true;
    }
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:DisableDeleteButton();">
<form name="piform" id="piform" runat="server">
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 500px">
    <tr>
        <td class="TITLE" colspan="8">
            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="100%" class="TITLE">
                        <b>
                            <asp:Label ID="lheader" runat="server" /></b>
                    </td>
                    <td width="100%" class="TITLE">
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td class="menuTD" align="right">
                <asp:Button ID="btnPrint" Text="Print" CssClass="all_button" OnClientClick="window.print();return false;" runat="server" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                <Columns>
                    <asp:TemplateField HeaderText="Item No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_code" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pack Key">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="pack_key" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Name">
                        <ControlStyle Width="150px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_name" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="pallet_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="batch_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Req. Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="dod_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Avail. Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="avail_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Hold Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="hold_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Picked Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="picked_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Bal. Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="itm_balance" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </td>
    </tr>
   <%-- <tr>
        <td id="tr_hold_title1" runat="server" visible="false" colspan="8" class="menuTD">&nbsp;</td>
    </tr>
    <tr id="tr_hold_title2" runat="server" visible="false">
        <td colspan="8" style="background-color:White">
            <br /><br /><br />
        </td>
    </tr>--%>
   <%-- <tr id="tr_hold_title3" runat="server" visible="false">
        <td colspan="8" class="TITLE" style="color:Blue">
            Please release items holding in CO
        </td>
    </tr>
    <tr id="tr_hold_gv" runat="server" visible="false">
        <td colspan="8">
            <asp:GridView ID="GridView2" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                <Columns>
                    <asp:TemplateField HeaderText="Item No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="cod_itm_code" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pack Key">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="cod_pack_key" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Name">
                        <ControlStyle Width="150px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_name" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="cod_pallet_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="cod_batch_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="CO Order Code">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="co_code" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Hold Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="hold_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
            <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
        </td>
    </tr>--%>
</table>
<asp:HiddenField ID="editMode" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
