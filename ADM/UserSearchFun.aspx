<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UserSearchFun.aspx.vb" Inherits="UserSearchFun" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Search Function</title>
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form method="post" name="searchform" id="searchform" onsubmit="" runat="server" defaultbutton="Submit">
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="757">
            <tr>
                <td colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lblTitle" runat="server"></asp:Label></b>
                            </td>
                            <td width="100%" class="TITLE">
                                <asp:Button ID="NewBtn" Text="New" runat="server" CssClass="all_button" />
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
                <td width="15%" class="LabelTD">
                    <font size="2">
                        <asp:Label ID="lblCode" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="usr_id" runat="server" Width="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD">
                    <font size="2"><font size="2">
                        <asp:Label ID="lblName" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="Usr_name" runat="server" Width="400" />
                </td>
            </tr>
            <!--
        **********************
        -->
            <tr>
                <td colspan="2">
                    <asp:Button ID="Submit" runat="server" Text="Search" class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <br />
    <br />
    <!--
        **********************
        Modify Here
        -->
    <div id="div2">
        <asp:GridView ID="gvrsList" runat="server" Height="10px" Width="100%" AllowPaging="True"
            Font-Names="Arial" Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False"
            EmptyDataText="No Record Found." BackColor="White" BorderColor="#CCCCCC" BorderStyle="None"
            BorderWidth="1px" CellPadding="3" CaptionAlign="Top" DataKeyNames="usr_id"  PageSize="20">
            <PagerStyle Font-Size="10px" ForeColor="#4C60B6" BackColor="#006699" HorizontalAlign="Left"
                BorderColor="Transparent" BorderWidth="1px" Font-Bold="True" Font-Names="Arial"
                Font-Strikeout="False" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="usr_id" DataNavigateUrlFormatString="UserMaster.aspx?usr_id={0}"
                    DataTextField="usr_id" HeaderText="User ID">
                    <ItemStyle Font-Size="11px" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="usr_name" HeaderText="User Name">
                    <ControlStyle Width="200px" />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="usr_status" HeaderText="Status">
                    <ControlStyle Width="200px" />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>                
                <asp:TemplateField ControlStyle-Width="50px">
                    <ItemTemplate>
                        <asp:Button ID="btnDelete" runat="server" Height="22px" Font-Size="11px" CommandName="Delete"
                            Text="Delete" CssClass="all_button" Font-Bold="false" />
                    </ItemTemplate>
                    <HeaderStyle Width="50px" />
                </asp:TemplateField>
            </Columns>
            <RowStyle BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px" ForeColor="Black"
                Font-Size="9px" />
            <HeaderStyle BackColor="#006699" BorderColor="White" BorderStyle="Inset" BorderWidth="1px"
                Font-Names="Arial" ForeColor="White" Font-Bold="True" Font-Size="11px" HorizontalAlign="Center"
                VerticalAlign="Middle" Wrap="True" />
            <FooterStyle BackColor="White" ForeColor="#000066" />
            <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
        </asp:GridView>
    </div>
    <!--
    **********************
    -->
    <asp:HiddenField ID="user_type" runat="server" />
    </form>
</body>
</html>
