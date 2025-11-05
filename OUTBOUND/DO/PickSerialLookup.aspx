<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PickSerialLookup.aspx.vb" Inherits="OUTBOUND_DO_PickSerialLookup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Pick List Lookup</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">
    function saveok() {
        document.luform.moduleAction.value = "SAVEOK";
        document.luform.submit();
    }
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form id="luform" runat="server">
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 700px">
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
                <asp:Button ID="saveBtn1" runat="server" Text="OK" CssClass="all_button" Visible="false" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td class="menuTD" align="right">&nbsp;</td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td class="LabelTD" nowrap width="90px">
            <font size="2">
                <asp:Label ID="lbl_ITM_CODE" runat="server" /></font>
        </td>
        <td nowrap width="100px">
            <font size="2">
                <asp:Label ID="ITM_CODE" runat="server" />
            </font>
        </td>
        <td class="LabelTD" nowrap width="90px">
            <font size="2">
                <asp:Label ID="lbl_ITM_SKU_NO" runat="server" /></font>
        </td>
        <td nowrap width="100px">
            <font size="2">
                <asp:Label ID="ITM_SKU_NO" runat="server" />
            </font>
        </td>
        <td class="LabelTD" nowrap width="90px">
            <font size="2">
                <asp:Label ID="lbl_PACK_KEY" runat="server" /></font>
        </td>
        <td nowrap width="60px">
            <font size="2">
                <asp:Label ID="PACK_KEY" runat="server" />
            </font>
        </td>
        <td class="LabelTD" nowrap width="90px">
            <font size="2">
                <asp:Label ID="lbl_REQ_QTY2" runat="server" /></font>
        </td>
        <td nowrap>
            <font size="2">
                <asp:Label ID="REQ_QTY2" runat="server" style="text-align:right" />
            </font>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                <Columns>
                    <asp:TemplateField HeaderText="">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Button ID="btnSelect" Text="Select" CommandName="SELECT" Font-Size="11px" CssClass="all_button" Font-Bold="false" runat="server" />
                            <asp:HiddenField ID="ILOC_SEQ" runat="server" />
                            <asp:HiddenField ID="ILBS_SEQ" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Stock Qty2">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="stock_qty2" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Bal Qty2">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="ilbs_qty2" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Serial No.">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="ilbs_serial_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Drum ID">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="ilbs_drum_id" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Drum Level">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="ilbs_drum_level" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_pallet_no" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">                        
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_batch_no" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Expiry Date">                        
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_expiry_date" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="WH">
                        <ControlStyle Width="25px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_wh" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Location">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_loc" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Floor">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_floor" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Area">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_area" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Rack">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_rack" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField> 
                    <asp:TemplateField HeaderText="Bin">
                        <ControlStyle Width="25px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_bin" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
            <asp:Button ID="saveBtn2" runat="server" Text="OK" CssClass="all_button" Visible="false" />
            <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
        </td>
    </tr>
</table>
<asp:HiddenField ID="IMP_CODE" runat="server" />
<asp:HiddenField ID="STORER_CODE" runat="server" />
<asp:HiddenField ID="CO_CODE" runat="server" />
<asp:HiddenField ID="DO_CODE" runat="server" />
<asp:HiddenField ID="PALLET_NO" runat="server" />
<asp:HiddenField ID="BATCH_NO" runat="server" />
<asp:HiddenField ID="PLD_SEQ" runat="server" />
<asp:HiddenField ID="editMode" runat="server" />
<asp:HiddenField ID="moduleAction" runat="server" />
<asp:HiddenField ID="ORG_LOC" runat="server" />
<asp:HiddenField ID="DOD_DISP_SEQ" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>

