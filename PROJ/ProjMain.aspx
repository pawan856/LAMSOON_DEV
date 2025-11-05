<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ProjMain.aspx.vb" Inherits="PROJ_ProjMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Project Master</title>
    <link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">
    
    <script language ="javascript" src="../js/validation.js"></script>
    <script language ="javascript" src="../js/JS_Calendar.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%">
            <tr>
                <td class="TITLE" colspan="6">
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
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button"  runat ="server" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_code" runat="server" />: </font>
                </td>
                <td width="30%">
                    <asp:Label ID="prj_code" runat="server" Text="" Font-Size="10"></asp:Label>
                </td>
                <td class="LabelTD" width="20%" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_status" runat="server" /></font>
                </td>
                <td width="30%">
                    <font size="2">
                        <asp:Label ID="prj_status" runat="server" /></font>
                </td>
            </tr>
           <tr>
                <td class="LabelTD" width="25%" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /> </font>
                </td>
                <td colspan="5">
                      <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_name" runat="server" /> </font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="prj_name" runat="server" Width="300" />
                </td>
            </tr>
            <tr>
             <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_date" runat="server" /></font>
                </td>
                <td>
                     <asp:TextBox ID="prj_date" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="prj_date" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                    </a>                    
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_end_date" runat="server" /></font>
                </td>
                <td colspan="5">
                     <asp:TextBox ID="prj_end_date" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="prj_end_date" PopupButtonID="btnDATE_ID2" Format="dd/MM/yyyy" />
                    </a>                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lprj_desc" runat="server" />:</font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="prj_desc" runat="server" TextMode="MultiLine" Width="400" Height="80" />
                </td>
            </tr>
            <tr>
                <td colspan="6" align="left" width="100%">
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
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
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
