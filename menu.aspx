<%@ Page Language="VB" AutoEventWireup="false" CodeFile="menu.aspx.vb" Inherits="menu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
   <meta http-equiv="content-type" content="text/html; charset=<%=G_HTML_LANG_ENCODING%>">
    <title>WMS v3.02</title>    
   
    <link rel="stylesheet" href="stylesheet/link.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">
    <style type="text/css">
        .style2
        {
            font-size: 16px;
            font-weight: bold;
        }
    </style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form id="menuform" runat="server">
<div style="height: 100px; width: 100%" runat="server" id="menuDiv">
        <asp:Panel runat="server" ID="menupan" BorderStyle="None" BorderWidth="0" ScrollBars="None"
            Width="100%">
            <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%;" style="background-color: white">
                <tr>
                    <td>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                            <tr>
                                <td width="50%" align="left" style="background-color: #FFFFFF; vertical-align: middle">
                                    <a href="main_news.aspx" target="main"> <img src="images/version/ae_right_banner.jpg" /></a>
                                </td>
                                <td valign="middle" align="right" style="background-color: #FFFFFF; vertical-align: middle" width="15%">
                                    <!--
                                    <asp:Button ID="btnInvisible" runat="server" BorderStyle="None" BackColor="Transparent" OnClientClick="return false;" />                                                   
                                    <asp:ImageButton ID="imgalert" runat="server" CssClass="image_btn" ImageUrl="~/images/email_yes.gif" AlternateText="Alerts" Style="cursor: pointer"/>
                                    -->
                                    <asp:Image runat="server" ID="imgDB" ImageUrl="" />
                                </td>
                                <td align="right" style="background-color: #FFFFFF; vertical-align: top" width="35%">
                                    <table border="0" cellspacing="0" cellpadding="0" align="right" bgcolor="#FFFFFF">
                                        <tr style="background-color: #FFFFFF">                                            
                                            <td style="background-color: #FFFFFF;">
                                               <asp:Image runat="server" ID="Image1" ImageUrl="~/images/version/naf_right_banner.png" />
                                            </td>
                                        </tr>
                                        <tr style="background-color: #FFFFFF">
                                            <td style="vertical-align: bottom; padding-top: 5px; background-color: #FFFFFF; vertical-align: top">
                                                <table border="0" cellpadding="0" cellspacing="0" align="right" style="background-color: #FFFFFF">
                                                    <tr style="background-color: #FFFFFF">
                                                       <%-- <td style="vertical-align: middle; font-size: 14px; background-color: #FFFFFF; font-weight:bold;" width="80%" align="left">
                                                            Selected DB: <asp:Label ID="l_conn" runat="server"></asp:Label>
                                                        </td> --%>
                                                        <td style="vertical-align: middle; font-size: 14px; background-color: #FFFFFF" width="80%" align="left">
                                                            <asp:Label ID="logined_user" runat="server"></asp:Label>
                                                        </td>                                                        
                                                        <td style="vertical-align: middle; font-size: 14px; height:35px; background-color: #FFFFFF" width="10%" nowrap>
                                                            <asp:ImageButton ID="btnLogout" runat="server" CssClass="image_btn" ImageUrl="~/images/logout2.png" AlternateText="Alerts" Style="cursor: pointer"/>                                                            
                                                        </td>                                                        
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        <asp:Repeater ID="rept" runat="server">
                            <HeaderTemplate>
                                <table cellspacing="1" cellpadding="1" width="100%" align="center" bgcolor="#cccccc"
                                    border="0">
                                    <tbody>
                                        <tr bgcolor="#FFFFFF">
                            </HeaderTemplate>
                            <ItemTemplate>
                                <td onmouseover="this.bgColor='#eeeeee'" onmouseout="this.bgColor=''" width="5%"
                                    nowrap class="menu_link">
                                    <div align="center">
                                        <asp:HyperLink runat="server" ID="menuLink" CssClass="link" Target="_self" />                                        
                                    </div>
                                </td>
                            </ItemTemplate>
                            <FooterTemplate>
                                </tr> </tbody> </table>
                            </FooterTemplate>
                        </asp:Repeater>
                    </td>
                </tr>
                <tr>
                    <td style="background-color: #FFFFFF">
                        &nbsp;
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </div>

</form>
<form name="urlGoFrom" method="post"></form>
<iframe height="0" width="0" src="blank.html" name = "changePageIframe" frameborder="1" style="position:absolute; top:0; left:0; width:0; height:0;"></iframe>
</body>
</html>
