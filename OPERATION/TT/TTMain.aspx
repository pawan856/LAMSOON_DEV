<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TTMain.aspx.vb" Inherits="TTMain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Search Function</title>
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form method="post" name="searchform" id="searchform" onsubmit="" runat="server" defaultbutton="Submit">
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="600">
            <tr>
                <td colspan="4" width="100%">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lblTitle" runat="server"></asp:Label></b>
                            </td>
                            <td width="100%" class="TITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <!--
        **********************
        Modify Here
        -->
            <tr>
                <td class="LabelTD">
                    <font size="2"><font size="2">
                        <asp:Label ID="lblTTType" runat="server"></asp:Label>:</font>
                </td>
                <td runat="server">
                    <asp:DropDownList ID="TTType" runat="server" AutoPostBack = "true">
                        <asp:ListItem Value="RO" Text = "RO"></asp:ListItem>
                        <asp:ListItem Value="GR" Text = "GR"></asp:ListItem>
                        <asp:ListItem Value="CO" Text = "CO"></asp:ListItem>
                        <asp:ListItem Value="DO" Text = "DO"></asp:ListItem>
                        <asp:ListItem Value="TRACKNO" Text = "Tracking No."></asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td class="LabelTD">
                    <font size="2"><font size="2">
                        <asp:Label ID="lblTTNo" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="TTNo" runat="server" Width="400" />
                </td>
            </tr>        
            <tr>
                <td colspan="4">
                    <asp:Button ID="Submit" runat="server" Text="Search" class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <br />
    <br />
    <table align="center" >
    <tr>
    
    <td class="TITLE"><asp:Label ID="lblsearchhd" runat ="server"></asp:Label></td>
    <td class="TITLE"></td>
    </tr>
    <tr>
    <td  valign="top">
<div id="div3" align="center">
        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="250" AllowPaging="True"
            Font-Names="Arial" Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
            CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" PageSize="50">
               <RowStyle CssClass="GV" />
                <PagerStyle Font-Size="10px" ForeColor="#4C60B6" BackColor="#006699" HorizontalAlign="Left"
                BorderColor="Transparent" BorderWidth="1px" Font-Bold="True" Font-Names="Arial"
                Font-Strikeout="False" />
                <Columns>
                <asp:TemplateField HeaderText="No.">
                    <HeaderStyle Width="40px" />
                    <ItemTemplate>
                        <asp:LinkButton ID="gv1NO" runat="server" Height="22px" Font-Size="11px" Font-Bold="false" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Tracking No.">
                    <HeaderStyle Width="100px" />
                    <ItemTemplate>
                        <asp:Label ID="gv1TRACKNO" runat="server" Height="22px" Font-Size="11px" Font-Bold="false" />
                    </ItemTemplate>
                </asp:TemplateField>                
                <asp:TemplateField HeaderText="Creation Date ">
                    <HeaderStyle Width="120px" />
                    <ItemTemplate>
                        <asp:Label ID="gv1CD" runat="server" Height="22px" Font-Size="11px" Font-Bold="false" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
        </asp:GridView>
    </div>
    </td>
    <td valign="top">
    <!--
        **********************
        Modify Here
        -->
    <div id="div2" align="center">
            <asp:GridView ID="gvrsList" runat="server" Height="10px" Width="200" AllowPaging="True"
            Font-Names="Arial" Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
            CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" PageSize="50">
               <RowStyle CssClass="GV" />
                <PagerStyle Font-Size="10px" ForeColor="#4C60B6" BackColor="#006699" HorizontalAlign="Left"
                BorderColor="Transparent" BorderWidth="1px" Font-Bold="True" Font-Names="Arial"
                Font-Strikeout="False" />
                <Columns>
                    <asp:TemplateField HeaderText="GR" ControlStyle-Width ="50px">
                        <ItemTemplate>
                            <asp:Label ID="TTGR" runat="server" Font-Size="11px"  Font-Bold="false" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>                
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:Image ID="TTROGR_SB" runat="server" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="RO"  ControlStyle-Width ="50px">
                        <ItemTemplate>
                            <asp:Label ID="TTRO" runat="server" Font-Size="11px" Font-Bold="false" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:Image ID="TTGRCO_SB" runat="server" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>               
                    <asp:TemplateField  HeaderText="CO" ControlStyle-Width ="50px">
                        <ItemTemplate>
                            <asp:Label ID="TTCO" runat="server" Font-Size="11px" Font-Bold="false" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:Image ID="TTCODO_SB" runat="server" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>                
                    <asp:TemplateField  HeaderText="DO" ControlStyle-Width ="50px">
                        <ItemTemplate>
                            <asp:Label ID="TTDO" runat="server" Font-Size="11px" Font-Bold="false" Height="18px" />
                        </ItemTemplate>
                    </asp:TemplateField>                
                </Columns>
        <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
        </asp:GridView>
    </div>
    </td>
    </tr>
    </table>
    <!--
    **********************
    -->
    </form>
</body>
</html>
