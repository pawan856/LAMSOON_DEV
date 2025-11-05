<%@ Page Language="VB" AutoEventWireup="false" CodeFile="storerMaster.aspx.vb" Inherits="MASTER_SM_storerMaster" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Storer Master</title>
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


function OpenFieldCustom() {

    removeAllElementFromForm(document.hiddenForm);
    window.open("", "FieldMast", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=920,height=600,left=5,top=15").focus();

    setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("STORER_CODE").innerHTML);    

    document.hiddenForm.action = "../FIELD_MASTER/FIELD_Main.aspx";
    document.hiddenForm.target = "FieldMast";
    document.hiddenForm.submit();
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form method="post" name="myform" id="myform" onsubmit="" runat="server">
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
                    &nbsp;&nbsp;
                    <asp:Button runat="server" ID="btnFieldMast" Text="Field Customization" CssClass="all_button" OnClientClick="javascript:OpenFieldCustom();return false;" Visible="false" />
                    &nbsp;&nbsp;

                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                    </td>
            </tr>
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:Label ID="STORER_CODE" runat="server" Text="" Font-Size="10" MaxLength="20"></asp:Label>
                </td>
            </tr>
            <tr>
               <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lbl_STO_STATUS" runat="server" />
                </td>
                <td colspan="7" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STO_STATUS" runat ="server" ></asp:DropDownList>
                    </font>
                </td> 
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_SHORTNAME" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_SHORTNAME" runat="server" Width="400px" MaxLength="20" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_NAME" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_NAME" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_NAME_CH" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_NAME_CH" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_ADDR" runat="server" />:</font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_ADDR1" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="STO_ADDR2" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="STO_ADDR3" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_COUNTRY" runat="server"  /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_COUNTRY" runat="server" MaxLength="50" />
                    </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_REGION" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_REGION" runat="server" MaxLength="50" />
                </td>
           </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_AREA" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="STO_AREA" runat="server" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_REM" runat="server" /></font>
                </td>
                <td colspan="7">
                    <%--<asp:TextBox ID="STO_REM" runat="server"  MaxLength="200" BackColor="YellowGreen"/>--%>
                    <asp:DropDownList ID="STO_REM" runat ="server"  BackColor="YellowGreen">
                        <asp:ListItem Value="">--Select--</asp:ListItem>
                        <asp:ListItem Value="NEW">By Rules and SO</asp:ListItem>
                        <asp:ListItem Value="NEW_NEW">By SO</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap style="border: 2px;">
                    <font size="2">
                        <asp:Label ID="lbl_STO_CONT_PER_GEN" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_CONT_PER_GEN" runat="server"  BackColor="YellowGreen" MaxLength="50"/>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_CONT_TEL_GEN" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_CONT_TEL_GEN" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_CONT_DEPT_GEN" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_CONT_DEPT_GEN" runat="server" MaxLength="30" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_CONT_EMAIL_GEN" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_CONT_EMAIL_GEN" runat="server" Width="165px" MaxLength="30" />
                </td>                
            </tr>
           <tr>
	            <td class="LabelTD" nowrap style="border: 2px;">
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_PER_MGR" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_PER_MGR" runat="server" MaxLength="50" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_TEL_MGR" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_TEL_MGR" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_DEPT_MGR" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_DEPT_MGR" runat="server" MaxLength="30" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_EMAIL_MGR" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_EMAIL_MGR" runat="server" Width="165px" MaxLength="30"  />
	            </td>                
            </tr>
            <tr>
	            <td class="LabelTD" nowrap style="border: 2px;">
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_PER_ACC" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_PER_ACC" runat="server" MaxLength="50" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_TEL_ACC" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_TEL_ACC" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_DEPT_ACC" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_DEPT_ACC" runat="server" MaxLength="30" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_EMAIL_ACC" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_EMAIL_ACC" runat="server" Width="165px" MaxLength="30" />
	            </td>                
            </tr>
            <tr>
	            <td class="LabelTD" nowrap style="border: 2px;">
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_PER_ORD" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_PER_ORD" runat="server" MaxLength="50" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_TEL_ORD" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_TEL_ORD" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_DEPT_ORD" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_DEPT_ORD" runat="server" MaxLength="30" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_EMAIL_ORD" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_EMAIL_ORD" runat="server" Width="165px" MaxLength="30" />
	            </td>                
            </tr>
            <tr>
	            <td class="LabelTD" nowrap style="border: 2px;">
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_PER_LOG" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_PER_LOG" runat="server" MaxLength="50" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_TEL_LOG" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_TEL_LOG" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_DEPT_LOG" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_DEPT_LOG" runat="server" MaxLength="30" />
	            </td>
	            <td class="LabelTD" nowrap>
	                <font size="2">
		            <asp:Label ID="lbl_STO_CONT_EMAIL_LOG" runat="server" /></font>
	            </td>
	            <td>
	                <asp:TextBox ID="STO_CONT_EMAIL_LOG" runat="server" Width="165px" MaxLength="30" />
	            </td>                
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_FAX" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_FAX" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_MAIN_TEL" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="STO_MAIN_TEL" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_WEBSITE" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="STO_WEBSITE" runat="server" Width="200px" MaxLength="30" />
                </td>
            </tr>
           <tr>
               <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_CURR" runat="server" /></font>
                </td>
                <td colspan="3" runat="server">
                    <asp:DropDownList ID="STO_CURR" runat ="server" ></asp:DropDownList>
                </td>
            </tr>            
            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STO_BATCH_FIELD_REF" runat="server" Text="Batch No Pattern" /></font>:
                </td>
                <td colspan="7" align="left">
                    <asp:RadioButtonList ID="STO_BATCH_FIELD_REF" runat="server" RepeatDirection="Horizontal"  />
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
    </form>
    <form id="hiddenForm" name="hiddenForm" method="post" />
</body>
</html>
