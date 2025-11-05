<%@ Page Language="VB" AutoEventWireup="false" CodeFile="locLookup.aspx.vb" Inherits="locLookup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language ="javascript" src="../js/validation.js"></script>
<script language ="javascript" src="../js/JS_Calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Search Function</title>
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">    
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form method="post" name="searchform" id="searchform" onsubmit="" runat="server" defaultbutton="Submit">
        
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="757">
            <tr>
                <td colspan="2">
                    <table cellspacing="0" cellpadding="0" border="0" style="border-width:0px;border-collapse:collapse;align:center;width:100%;">
                    <tr>
				        <td class="TITLE" style="width:100%;"><b><asp:Label ID="lblTitle" runat="server"></asp:Label></b></td>
			        </tr>
		        </table>
                </td>
            </tr>
            <!--
        **********************
        Modify Here
        -->
        <asp:Label ID="testlabel" runat="server" >
        </asp:Label>
        <asp:HiddenField ID="pForm" runat ="server" />
        <asp:HiddenField ID="pItemList" runat ="server" />
        <asp:HiddenField ID="pFunc" runat ="server" />
        <asp:HiddenField ID="wh" runat ="server" />
        <asp:HiddenField ID="mwh" runat ="server" />
        <asp:HiddenField ID="sc" runat ="server" />
        <asp:HiddenField ID="itemCode" runat ="server" />
        <asp:HiddenField ID="packkey" runat ="server" />
        <asp:HiddenField ID="WH_CODE" runat ="server" />
        <asp:HiddenField ID="FL_NUM" runat ="server" />
        <asp:HiddenField ID="AR_CODE" runat ="server" />
        <asp:HiddenField ID="RK_CODE" runat ="server" />
        <asp:HiddenField ID="BN_CODE" runat ="server" />
        <asp:HiddenField ID="FUN_CODE" runat ="server" />
            <asp:TableRow ID="trWH_CODE" runat ="server">
            </asp:TableRow>
            <asp:TableRow ID="trFL_NUM" runat ="server">
            </asp:TableRow>            
            <asp:TableRow ID="trAR_CODE" runat ="server">
            </asp:TableRow>
            <asp:TableRow ID="trRK_CODE" runat ="server">
            </asp:TableRow>
            <asp:TableRow ID="trBN_CODE" runat ="server">
            </asp:TableRow>            
            <!--
        **********************
        -->
            <tr>
                <td colspan="2">
                    <asp:Button ID="Submit" runat="server" Text="Select" Visible = false class="all_button" UseSubmitBehavior = false />
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

    <!--
    **********************
    -->
    </form>
</body>
</html>
