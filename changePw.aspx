<%@ Page Language="VB" AutoEventWireup="false" CodeFile="changePW.aspx.vb" Inherits="changePW" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language ="javascript" src="js/validation.js"></script>
<script language ="javascript" src="js/JS_Calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Search Function</title>
    <link rel="stylesheet" href="stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="stylesheet/general.css" type="text/css">    
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="document.forms[0].Oldpw.focus();">
    <br />
    <form method="post" name="searchform" id="searchform" onsubmit="" runat="server" defaultbutton="Submit">
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="60%">
            <tr>
                <td colspan="2" class="TITLE">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lblTitle" runat="server"></asp:Label></b>
                            </td>
                            <td width="100%" class="TITLE">
                                &nbsp;
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
                <td class="LabelTD" width="220px">
                    <font size="2">
                        <asp:Label ID="lblOldpw" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="Oldpw" runat="server" Width="80px" TextMode ="Password" 
                        MaxLength="10" />
                </td>
            </tr>        
            <tr>
                <td class="LabelTD" width="220px">
                    <font size="2">
                        <asp:Label ID="lblNewpw" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="Newpw" runat="server" Width="80px" TextMode ="Password" MaxLength="10" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" width="220px">
                    <font size="2">
                        <asp:Label ID="lblConfirmpw" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:TextBox ID="Confirmpw" runat="server" Width="80px" TextMode ="Password" MaxLength="10" />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Button ID="Submit" runat="server" Text="Save" class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <br />
    <br />
    </form>
</body>
</html>
