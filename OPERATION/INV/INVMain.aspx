<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INVMain.aspx.vb" Inherits="OPERATION_INV_INVMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <title>Invoice</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>

<script language="javascript">

function maskKeyAndCheck(objEvent, loc) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    
    if (loc != null) {
        if (loc != '') {
            if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) {
                return true;
            } else {
                return false;
            }            
        } else {
            alert("You Must Select Location First!");
            return false;
        }

    } else {
        alert("You Must Select Location First!");
        
        return false;
    }
}

function PrintInvoice() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "invPrintout", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,resizable=yes,width=1030,height=700,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.h_STORER_CODE.value);

        setInterfaceDataToForm(document.hiddenForm, "listmode", document.myform.h_listmode.value);
        setInterfaceDataToForm(document.hiddenForm, "INVH_NO", document.myform.h_INVH_NO.value);
        document.hiddenForm.action = "INV_PRINTOUT/inv_printout.aspx";
        document.hiddenForm.target = "invPrintout";
        document.hiddenForm.submit();
}

function calAmt(RateId, Qtyid, Amtid, HDid) {
    var oRate = eval("document.getElementById('" + RateId + "')").value;
    var oQty = eval("document.getElementById('" + Qtyid + "')").value;
    
    eval("document.getElementById('" + Amtid + "')").innerHTML = oRate * oQty;
    eval("document.getElementById('" + HDid + "')").value = oRate * oQty;
    
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />    
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                </td>
            </tr>
            <asp:HiddenField ID="lheader" runat ="server" />
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=OP_INV'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <!--<input id="btnPrint" type="button" value="Print Invoice"
                            onclick="Javascript:PrintInvoice();"
                            class="all_button"  />-->
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_INVH_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="INVH_NO" runat="server" />
                        <asp:HiddenField id="h_STORER_CODE" runat ="server" />
                        <asp:HiddenField id="h_INVH_NO" runat ="server" />
                        <asp:HiddenField id="h_listmode" runat ="server" />
                    </font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack="true"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_INVH_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="INVH_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="img_INVH_DATE" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="INVH_DATE" PopupButtonID="img_INVH_DATE" Format="dd/MM/yyyy" />  
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_INVH_FR_MONTH" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="INVH_FR_MONTH" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="img_INVH_FR_MONTH" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="INVH_FR_MONTH" PopupButtonID="img_INVH_FR_MONTH" Format="dd/MM/yyyy" />   
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_INVH_TO_MONTH" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="INVH_TO_MONTH" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="img_INVH_TO_MONTH" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="INVH_TO_MONTH" PopupButtonID="img_INVH_TO_MONTH" Format="dd/MM/yyyy" />         
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_INVH_GEN_BY" runat="server" /></font>
                </td>
                <td nowrap colspan="3">
                    <font size="2">
                        <asp:TextBox ID="INVH_GEN_BY" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>                
            </tr>
            <!-- <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_CUS_CODE" runat="server" /></font>
                </td>
                <td>
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                    <td id="Td2" width="50px" runat="server"><asp:DropDownList ID="CUS_CODE" runat="server" MaxLength="20" 
                        AutoPostBack="True" CssClass="REQUIRED"></asp:DropDownList></td>
                    <td><font size="2"></font></td>
                    </tr>
                    </table>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_CUS_NAME" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                    <asp:textbox ID="CUS_NAME" runat="server" ReadOnly="true" 
                        BackColor="#dddddd" Width="200px" /></font>
                </td>
            </tr> -->
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lbl_INVH_ADDR" runat="server" />
                </td>
                <td colspan="3">
                        <asp:TextBox ID="INVH_ADDR1" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                    <br />
                        <asp:TextBox ID="INVH_ADDR2" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                    <br />
                        <asp:TextBox ID="INVH_ADDR3" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVH_ATTN" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:TextBox ID="INVH_ATTN" runat="server" MaxLength="50" Width="165px"></asp:TextBox>
                    </font>
                </td>
            </tr> 
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVH_TEL" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="INVH_TEL" runat="server" MaxLength="20" 
                        Width="165px" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                 <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVH_FAX" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="INVH_FAX" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                </td>
            </tr>       
            <tr>
                <td colspan="4">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td colspan="4" align="left" width="100%">
                    <table border="1" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_cb" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cb" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lub" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lub" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_cd" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cd" runat="server" style="white-space:nowrap" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lud" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lud" runat="server" style="white-space:nowrap" />
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
              <tr>
                <td class="TITLE" colspan="4">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b><asp:Label ID="lbl_ImageHd" runat="server" /></b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">              
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                        <td class="menuTD" align="left">
                             <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                            <asp:Button ID="genInvBtn" runat="server" Text="Generate Invoice" CssClass="all_button" UseSubmitBehavior ="false" />
                        </td>
                        <td class="menuTD" align="right">
                           <asp:RadioButtonList runat="server" ID="listmode" RepeatColumns="3" 
                                RepeatLayout="Flow" AutoPostBack="true">
                               <asp:ListItem Selected="True" Value="DEFAULT">Show Summary</asp:ListItem>
                               <asp:ListItem Value="ITEM">Show Items</asp:ListItem>
                               <asp:ListItem Value="LOCATION">Show Location</asp:ListItem>
                           </asp:RadioButtonList>
                        </td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="invd_seq" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                              <asp:TemplateField HeaderText="No.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_display_seq" runat="server" Font-Size="11px" style="text-align:center" Width="30px" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Code">
                                <ControlStyle Width="120px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_itm_code" runat="server" Font-Size="11px" Width="120px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_pack_key" runat="server" Font-Size="11px" Width="60px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Location">
                                <ControlStyle Width="100px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_loc" runat="server" Font-Size="11px" Width="100px" MaxLength="100"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                                <asp:TemplateField HeaderText="Date">
                                    <ControlStyle Width="100px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="invd_tx_date" runat="server" Font-Size="11px" Width="100px" MaxLength="50"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                            <asp:TemplateField HeaderText="Type">
                                <ControlStyle Width="60px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_type" runat="server" Font-Size="11px" Width="60px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description">
                                <ControlStyle Width="200px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_desc" runat="server" Font-Size="11px" Width="200px" MaxLength="100"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Unit">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_unit" width="80px" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Currency">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="invd_curr" runat ="server" Font-Size="11px"></asp:DropDownList>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Rate">
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_rate" width="80px" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ItemTemplate>
                                    <asp:TextBox ID="invd_qty" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Amount">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_invd_amount" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                    <asp:HiddenField id="invd_amount" runat ="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                        </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=OP_INV'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post"></form>
</body>
</html>