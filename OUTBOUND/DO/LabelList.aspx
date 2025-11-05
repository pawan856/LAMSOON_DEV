<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LabelList.aspx.vb" Inherits="OUTBOUND_DO_LabelList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Item Labels</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">
function saveok()
{
    document.luform.moduleAction.value = "SAVEOK";
    document.luform.submit();
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form id="luform" runat="server">
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 500px">
   <tr>
        <td class="TITLE" colspan="6">
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
        <td colspan="6" class="menuTD">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td class="menuTD" align="right">
                    <asp:Button ID="btnPrint" runat="server" Text="Print Label" cssclass="all_button" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="6">
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" >
                <Columns>
                    <asp:TemplateField HeaderText="Seq No.">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_disp_seq" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Code">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_itm_code" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pack Key">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_pack_key" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Name">
                        <ControlStyle Width="200px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_itm_desc" width="80" MaxLength="100" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true"/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Tracking No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_track_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="70px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_pallet_no" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Carton No.">
                        <ControlStyle Width="70px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_carton_no" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Job No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dod_ref_no" width="70" MaxLength="12" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty">
                        <ControlStyle Width="45px" />
                        <HeaderStyle HorizontalAlign="Right" Wrap="false" />
                        <ItemStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="label_qty" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                     <asp:TemplateField ControlStyle-Width="50px">
                            <ItemTemplate>
                                <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px"
                                    CommandName="Update" Text="Split" CssClass="all_button" Font-Bold="false" />
                                    <tr>
                                    <td colspan="11" nowarp>
                                    <div id="div2" style="display:none" runat="server">
                                    <asp:GridView ID="GridView2" runat="server" Height="10px" Width="100%" 
                                    Font-Names="Arial" Font-Overline="False" Font-Size="10px" BackColor="White" 
                                    BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left"
                                     AutoGenerateColumns="False" ShowFooter="false" ShowHeader="false" 
                                     OnRowDataBound="GridView2_RowDataBound">
                                     <Columns>             
                                         <asp:TemplateField HeaderText="Seq No.">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_disp_seq" width="28" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Item Code">
                                            <ControlStyle Width="80px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_itm_code" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Pack Key">
                                            <ControlStyle Width="40px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_pack_key" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Item Name">
                                            <ControlStyle Width="200px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_itm_desc" width="80" MaxLength="100" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true"/>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Tracking No.">
                                            <ControlStyle Width="100px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_track_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Pallet No.">
                                            <ControlStyle Width="70px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_pallet_no" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Carton No.">
                                            <ControlStyle Width="70px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_carton_no" width="70" MaxLength="20" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Job No.">
                                            <ControlStyle Width="80px" />
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="dod_ref_no" width="70" MaxLength="12" runat="server" Font-Size="11px" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Qty">
                                            <ControlStyle Width="45px" />
                                            <HeaderStyle HorizontalAlign="Right" Wrap="false" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="label_qty" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField ControlStyle-Width="48px">
                                            <ItemTemplate>
                                                <asp:label ID="empty_col" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                             </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                    </asp:GridView>
                                    </div>
                                 </td>
                            </tr>
                            </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                    </asp:TemplateField>	
                </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr>
        <td colspan="6" class="menuTD">
            <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
        </td>
    </tr>
</table>
<asp:HiddenField ID="IMP_CODE" runat="server" />
<asp:HiddenField ID="STORER_CODE" runat="server" />
<asp:HiddenField ID="PALLET_NO" runat="server" />
<asp:HiddenField ID="editMode" runat="server" />
<asp:HiddenField ID="moduleAction" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
