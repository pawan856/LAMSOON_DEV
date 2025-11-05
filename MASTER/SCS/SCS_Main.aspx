<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SCS_Main.aspx.vb" Inherits="OPERATION_SCS_Main" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>System Code Maintenance</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">



function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}


</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form method="post" name="myform" id="myform" onsubmit="" runat="server">
    <input type="hidden" name="moduleAction" value=""/>
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 900px">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                    <asp:label ID="lheader" runat ="server" />
                 </td>
            </tr>            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" UseSubmitBehavior="false" CssClass="all_button" />
                        <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" /> 
                       <asp:Button ID="DelBtn1" runat="server" Text ="Delete" CssClass="all_button" UseSubmitBehavior="False" /> 
                       </td>   
                </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_code" runat="server" />:</font>
                </td>
                <td runat="server">
                    <font size="2">
                         <asp:textbox ID="colc_code" runat="server" MaxLength="50" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_status" runat="server" />:</font></td>
                <td>
                    <font size="2">
                       <asp:DropDownList ID="Status" runat="server">
                           <asp:ListItem Value="">SELECT</asp:ListItem>
                           <asp:ListItem Value="A">ACTIVE</asp:ListItem>
                           <asp:ListItem Value="I">INACTIVE</asp:ListItem>
                    </asp:DropDownList>
                    </font>
                </td>
             </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_code_type" runat="server" />:</font></td>
                <td runat="server">
                    <font size="2">
                        <asp:DropDownList ID="colc_tabcol" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                <font size="2">
                    <asp:Label ID="lbl_display_seq" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:Textbox ID="colc_display_seq" runat="server" MaxLength="5" Width="50" />
                </td>
              </tr>
              <tr>
                    <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                    <asp:Label ID="lbl_colc_eng_value" runat="server" />:</font></td>
                    <td colspan="3"><font size="2" runat="server">
                        <asp:textbox ID="colc_eng_value" runat="server" maxlength="50" width="400" /></font>
                    </td>
              </tr>
             <tr>
                <td colspan="4" class="TITLE">
                    &nbsp;
                 </td>
            </tr>  
        </table>
    </div>
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" method="post" />
  </body>
</html>
