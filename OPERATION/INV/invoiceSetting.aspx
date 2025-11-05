<%@ Page Language="VB" AutoEventWireup="false" CodeFile="invoiceSetting.aspx.vb" Inherits="INV_invoiceSetting" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>UOM Maintenance</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">
function DisableDeleteButton() {
    var grp = document.getElementsByName("btnDelete");
    var count;
    count = grp.length;
    if (count == 1) {
       grp[0].disabled = true;
    }
}

function maskKey(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
    return false;
}

function maskDate(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
    return false;
}

function LocLookUp(lb_id) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"  marginheight="0">
    <br />
    <form id="myform"  runat="server">
    <asp:ToolkitScriptManager ID="ScriptManager1" runat="server">
    </asp:ToolkitScriptManager> 
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%">
            <tr>
                <td class="TITLE" colspan="2">
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
                <td colspan="2" class="menuTD">
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button"  OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" /></td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td id="Td1" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack="true" />
                    </font>
                </td>
            </tr>
            <tr>
               <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lbl_INVR_STATUS" runat="server" />
                </td>
                <td id="Td3" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="INVR_STATUS" runat ="server" ></asp:DropDownList>
                    </font>
                </td> 
            </tr>
            <asp:HiddenField ID="INV_SEQ" runat="server" Value="" />
            <tr>
                <td class="LabelTD" width="25%" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INV_DISPLAY_SEQ" runat="server" /> </font>
                </td>
                <td>
                    <asp:textbox ID="INV_DISPLAY_SEQ" runat="server" Text="" Font-Size="10" MaxLength="10" 
                        Width="110px"></asp:textbox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" width="25%" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_TYPE" runat="server" /> </font>
                </td>
                <td>
                <asp:UpdatePanel runat="server" id="udp1" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:DropDownList ID="INVR_TYPE" runat="server" AutoPostBack="true" />     
                    </ContentTemplate>
                </asp:UpdatePanel>
                    
                </td>
            </tr>

            <tr>
                <td class="LabelTD" width="25%" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_LOC" runat="server" /> </font>
                </td>
                <td>
                    <asp:textbox ID="INVR_LOC" runat="server" Text="" Font-Size="10" width="195px" MaxLength="20"></asp:textbox>&nbsp;
                    <asp:Image ID="Image_INVR_LOC" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_DESC" runat="server" /> </font>
                </td>
                <td>
                  <asp:UpdatePanel runat="server" ID="UDP2" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:textbox ID="INVR_DESC" runat="server" Text="" Font-Size="10" width="300px" MaxLength="100"></asp:textbox>    
                    </ContentTemplate>
                  </asp:UpdatePanel>                    
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_EFF_DATE" runat="server" /> </font>
                </td>
                <td>
                    <asp:TextBox ID="INVR_EFF_DATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="INVR_EFF_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_EXP_DATE" runat="server" /> </font>
                </td>
                <td>
                    <asp:TextBox ID="INVR_EXP_DATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                     <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="INVR_EXP_DATE" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_UNIT" runat="server" /> </font>
                </td>
                <td>
                    <asp:textbox ID="INVR_UNIT" runat="server" Text="" Font-Size="10" width="195px" MaxLength="20"></asp:textbox>
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_RATE" runat="server" /> </font>
                </td>
                <td>
                    <asp:textbox ID="INVR_RATE" runat="server" Text="" Font-Size="10" width="195px" MaxLength="14"></asp:textbox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_INVR_CURR" runat="server" /></font>
                </td>
                <td id="Td2" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="INVR_CURR" runat="server" />
                    </font>
                </td>
            </tr>            
            <tr>
                <td colspan="2" align="left" width="100%">
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
                                    <asp:Label ID="sys_cd" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lud" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lud" runat="server" />
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack1" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'" class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>