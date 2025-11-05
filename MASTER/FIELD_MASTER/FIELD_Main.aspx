<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FIELD_Main.aspx.vb" Inherits="OPERATION_FIELD_Main" %>


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
                        <input type="button" value="Close" class="all_button" onclick="javascript:window.close();" />
                       </td>   
                </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" Text="Storer Name" />:</font>
                </td>
                <td colspan="3">
                    <font size="2">
                         <asp:Label runat="server" ID="STORER_NAME" />
                    </font>
                </td>                
                
             </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_FUN_CODE" runat="server" text="Function Name" />:</font></td>
                <td>
                    <font size="2">
                        <asp:DropDownList ID="FUN_CODE" runat="server" AutoPostBack="true" >
                            <asp:ListItem Value="MAST_IM" Text="Item Master" />
                            <asp:ListItem Value="IB_RO" Text="Replenishment Order" />
                            <asp:ListItem Value="IB_RO_C" Text="Replenishment Order (for storer)" />
                            <asp:ListItem Value="IB_RO_RCP" Text="RO Print Order" />
                            <asp:ListItem Value="IB_GR" Text="Goods Receive" />
                            <asp:ListItem Value="IB_GR_RCP" Text="GR: Stock Receipt" />
                            <asp:ListItem Value="IB_GR_CRP" Text="GR: Cargo Receipt" />
                            <asp:ListItem Value="OB_CO" Text="Customer Order" />
                            <asp:ListItem Value="OB_CO_C" Text="Customer Order (for storer)" />                            
                            <asp:ListItem Value="OB_DO" Text="Delivery Order" />                          
                            <asp:ListItem Value="OB_DO_DN" Text="Delivery Note" />        
                            <asp:ListItem Value="INQ_001" Text="Location Stock Balance" />
                            <asp:ListItem Value="RPT_SB" Text="Stock Balance Printout (Report)" />
                        </asp:DropDownList>
                    </font>
                </td>                
                <td colspan="2">
                    <asp:Button runat="server" ID="btnSearch" Text="Search" CssClass="all_button" Visible="false" />
                </td>
              </tr>
              <tr>
                    <td colspan="4">
                        <asp:DataList runat="server" ID="field_list" width="100%" RepeatColumns="3" RepeatDirection="Horizontal">
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                            <ItemTemplate>
                                <div>
                                    <asp:CheckBox runat="server" ID="FLDO_FIELD_OPTION" />
                                    <asp:Label runat="server" ID="FLDO_FIELD_DESC" />
                                    <asp:hiddenfield runat="server" ID="FLDO_FIELD_NAME" />
                                </div>
                            </ItemTemplate>
                        </asp:DataList>
                    </td>                    
               </tr>
               <tr id="manTr" runat="server">
                    <td colspan="4">
                    <asp:Label runat="server" ID="manLBL" Text="(The Following Field is mandatory field or system use and should not be modified.)" /><br />
                       <asp:DataList runat="server" ID="exclude_list" width="100%" RepeatColumns="3" RepeatDirection="Horizontal">
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                            <ItemTemplate>
                                <div>
                                    <asp:CheckBox runat="server" ID="FLDO_FIELD_OPTION" Enabled="false" />
                                    <asp:Label runat="server" ID="FLDO_FIELD_DESC" />
                                    <asp:hiddenfield runat="server" ID="FLDO_FIELD_NAME" />
                                </div>
                            </ItemTemplate>
                       </asp:DataList>
                    </td>
               </tr>
               <tr id="rptTr" runat="server" visible="false">
                    <td colspan="4">
                    <asp:Label runat="server" ID="rptLbl" Text="(The Following fields will only show in download report.)" /><br />
                       <asp:DataList runat="server" ID="report_list" width="100%" RepeatColumns="3" RepeatDirection="Horizontal">
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                            <ItemTemplate>
                                <div>
                                    <asp:CheckBox runat="server" ID="FLDO_FIELD_OPTION" />
                                    <asp:Label runat="server" ID="FLDO_FIELD_DESC" />
                                    <asp:hiddenfield runat="server" ID="FLDO_FIELD_NAME" />
                                </div>
                            </ItemTemplate>
                       </asp:DataList>
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
