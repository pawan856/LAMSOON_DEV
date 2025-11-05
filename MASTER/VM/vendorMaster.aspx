<%@ Page Language="VB" AutoEventWireup="false" CodeFile="vendorMaster.aspx.vb" Inherits="vendorMaster" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Vendor Master</title>
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

function checkChange(obj) {
    if (obj.checked) {
        document.myform.VND_BLACKLIST_REASON.style.backgroundColor = "transparent";
        document.myform.VND_BLACKLIST_REASON.disabled = false;
    } else {
        document.myform.VND_BLACKLIST_REASON.style.backgroundColor = "#DDDDDD";
        document.myform.VND_BLACKLIST_REASON.disabled = true;
    }
}

function checkReasonCtrl() {
    if (document.myform.VND_BLACKLIST.checked) {
        //document.myform.VND_BLACKLIST_REASON.style.backgroundColor = "transparent";
        //document.myform.VND_BLACKLIST_REASON.disabled = false;
    } else {
        document.myform.VND_BLACKLIST_REASON.style.backgroundColor = "#DDDDDD";
        document.myform.VND_BLACKLIST_REASON.disabled = true;
    }
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="checkReasonCtrl();">
    <br />
    <form method="post" name="myform" id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%">
            <tr>
                <td class="TITLE" colspan="8">
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
                <td colspan="8" class="menuTD">
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" /></td>
            </tr>
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CODE" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_CODE" runat="server" Width="150px" MaxLength="20"></asp:TextBox>
                    <asp:FilteredTextBoxExtender runat="server" FilterType="Numbers, LowercaseLetters, UppercaseLetters, Custom" ValidChars="_ " TargetControlID="VND_CODE" />
                    &nbsp;&nbsp;&nbsp;<asp:Label ID="VND_CODE_HELP" runat="server" Text="Please input English characters and '_' only for the Vendor Code." ForeColor="Blue" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                    <font size="2">
                        <asp:Label ID="lSTORER_CODE" runat="server" /> </font>
                </td>
                <td colspan="7" runat="server">
                    <asp:DropDownList ID="STORER_CODE" runat ="server" ></asp:DropDownList>
                </td>
            </tr>
            <tr>
               <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lVND_STATUS" runat="server" />
                </td>
                <td colspan="7" runat="server">
                    <font size="2">
                            <asp:DropDownList ID="VND_STATUS" runat ="server" ></asp:DropDownList>
                </td> 
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_SHORTNAME" runat="server" />&nbsp;</font></td>
                <td colspan="7">
                    <asp:TextBox ID="VND_SHORTNAME" runat="server" Width="400px" MaxLength="20" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_NAME" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_NAME" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_NAME_CH" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_NAME_CH" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_ADDR" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_ADDR1" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="VND_ADDR2" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="VND_ADDR3" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_COUNTRY" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_COUNTRY" runat="server" MaxLength="50" />
                    </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_REGION" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_REGION" runat="server" MaxLength="50" />
                </td>
           </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_AREA" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_AREA" runat="server" MaxLength="50" />
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap style="border: 2px;">
                    <font size="2">
                        <asp:Label ID="lVND_CONT_PER1" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_PER1" runat="server" MaxLength="50" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_TEL1" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_TEL1" runat="server" onkeypress="return maskKey(event);"  MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_DEPT1" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_DEPT1" runat="server" MaxLength="30" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_EMAIL1" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_EMAIL1" runat="server" Width="165px" MaxLength="30" />
                </td>                
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_PER2" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_PER2" runat="server" MaxLength="50" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_TEL2" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_TEL2" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_DEPT2" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_DEPT2" runat="server"  MaxLength="30" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CONT_EMAIL2" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_CONT_EMAIL2" runat="server" Width="165px" MaxLength="30" />
                </td>                
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_FAX" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_FAX" runat="server" onkeypress="return maskKey(event);"  MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_MAIN_TEL" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_MAIN_TEL" runat="server" onkeypress="return maskKey(event);"  MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_WEBSITE" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="VND_WEBSITE" runat="server" Width="200px"  MaxLength="30" />
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_PROD_CAT" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_PROD_CAT" runat="server" Width="185px" MaxLength="20" />
                </td>
            </tr>            
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_BLACKLIST" runat="server" /></font>
                </td>
                <td colspan="7" valign="top">
                    <asp:CheckBox ID="VND_BLACKLIST" runat="server" Text="Yes/No" Height="50px" onclick="Javascript:checkChange(this);" />&nbsp;
                    <asp:TextBox ID="VND_BLACKLIST_REASON" runat="server" TextMode="MultiLine" Width="400" Height="80"  MaxLength="100"></asp:TextBox>
                </td>
            </tr>
          <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_PERF_QC" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="VND_PERF_QC" runat="server" MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_PERF_DELIV" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="VND_PERF_DELIV" runat="server" MaxLength="20" />
                </td>                
            </tr>                 
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_CURR" runat="server" /></font>
                </td>
                <td runat="server">
                    <asp:DropDownList ID="VND_CURR" runat ="server" ></asp:DropDownList>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_PAYTERM" runat="server" /></font>
                </td>
                <td colspan="5" runat="server">
                    <asp:DropDownList ID="VND_PAYTERM" runat ="server" ></asp:DropDownList>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lVND_REM" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="VND_REM" runat="server" TextMode="MultiLine" Width="400" Height="80" MaxLength="200" />
                </td>
            </tr>
            <tr>
                <td colspan="8" align="left" width="100%">
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
                <td colspan="8" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div id="div2">
    </div>
    </form>
</body>
</html>
