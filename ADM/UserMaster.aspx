<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UserMaster.aspx.vb" Inherits="UserMaster" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Search Function</title>
    <link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">
    
    <script language ="javascript" src="../js/validation.js"></script>
    <script language ="javascript" src="../js/JS_Calendar.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%">
            <tr>
                <td class="TITLE" colspan="4">
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
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../cms_search.aspx?menu_code=<%=fun_code%>'"
                        class="all_button" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_id" runat="server" />: 
                </td>
                <td>
                    <asp:textbox ID="usr_id" runat="server" Text="" onBlur="javascript:this.value = this.value.toUpperCase();" ></asp:textbox>
                </td>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_status" runat="server" />:
                </td>
                <td>
                    <asp:DropDownList ID="usr_status" runat ="server" ></asp:DropDownList>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_pwd" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_pwd" runat="server" TextMode ="Password" />
                </td>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_group" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_group" runat="server" MaxLength ="20"/>
                </td>                
            </tr>
            <tr id="tr_cus_code" runat="server">
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_cus_code" runat="server" />:
                </td>
                <td>
                    <asp:DropDownList ID="usr_cus_code" runat ="server" ></asp:DropDownList>
                </td>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_manager" runat="server" />:
                </td>
                <td>
                    <asp:DropDownList ID="usr_manager" runat ="server" ></asp:DropDownList>
                </td>                
            </tr>
            <tr id="tr_storer_code" runat="server" visible="false">
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_storer_code" runat="server" />:
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="usr_storer_code" runat ="server" ></asp:DropDownList>
                </td>          
            </tr>
            <tr id="tr_pref_storer" runat="server">
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_pref_storer" runat="server" />:
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="usr_pref_storer" runat ="server" ></asp:DropDownList>
                </td>          
            </tr>
            <tr id="tr_vnd_no" runat="server">
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_vnd_no" runat="server" />:
                </td>
                <td colspan="3">
                    <asp:DropDownList ID="usr_vnd_no" runat ="server" ></asp:DropDownList>
                </td>          
            </tr>
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_dob" runat="server" />: 
                </td>
                <td>
                    <asp:TextBox ID="usr_dob" runat="server" Width ="80"></asp:TextBox>
                     <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="usr_dob" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />        
                </td>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_start_date" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_start_date" runat="server" Width ="80"></asp:TextBox>
                     <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="usr_start_date" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />         
                </td>      
            </tr>
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_last_salary" runat="server" />: 
                </td>
                <td>
                    <asp:textbox ID="usr_last_salary" runat="server" Text=""></asp:textbox>
                </td>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_current_salary" runat="server" />:
                </td>
                <td>
                    <asp:textbox ID="usr_current_salary" runat="server" Text="" ></asp:textbox>
                </td>                
            </tr>   
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_fname" runat="server" />: 
                </td>
                <td>
                    <asp:textbox ID="usr_fname" runat="server" Text=""></asp:textbox>
                </td>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_sname" runat="server" />:
                </td>
                <td>
                    <asp:textbox ID="usr_sname" runat="server" Text="" ></asp:textbox>
                </td>                
            </tr>                        
            <tr>
                <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lusr_nickname" runat="server" />: 
                </td>
                <td colspan ="3">
                    <asp:textbox ID="usr_nickname" runat="server" Text=""></asp:textbox>
                </td>
            </tr>                                    
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_addr" runat="server" />:
                </td>
                <td colspan ="3">
                    <asp:TextBox ID="usr_addr1" runat="server" Width="400px" Height="22px" /><br />
                    <asp:TextBox ID="usr_addr2" runat="server" Width="400" /><br />
                    <asp:TextBox ID="usr_addr3" runat="server" Width="400" />
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_country" runat="server" />:
                </td>
                <td colspan ="3">
                    <asp:TextBox ID="usr_country" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_region" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_region" runat="server" />
                </td>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_area" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_area" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_home_tel" runat="server" />:
                </td>
                <td>
                    <asp:TextBox ID="usr_home_tel" runat="server" Width="100" />
                </td>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_mobile_tel" runat="server" />:
                </td>
                <td >
                    <asp:TextBox ID="usr_mobile_tel" runat="server" Width="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_email" runat="server" />:
                </td>
                <td coslpan="3">
                    <asp:TextBox ID="usr_email" runat="server" Width="400" />
                </td>
            </tr>
            <tr>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_title" runat="server" />:
                </td>
                <td >
                    <asp:TextBox ID="usr_title" runat="server" Width="200" />
                </td>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_work_title" runat="server" />:
                </td>
                <td >
                    <asp:TextBox ID="usr_work_title" runat="server" Width="200" />
                </td>
            </tr>          
            <tr>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_department" runat="server" />:
                </td>
                <td >
                    <asp:TextBox ID="usr_department" runat="server" Width="200" />
                </td>
                 <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_responsibility" runat="server" />:
                </td>
                <td >
                    <asp:TextBox ID="usr_responsibility" runat="server" Width="200" />
                </td>
            </tr> 
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lusr_remarks" runat="server" />:
                </td>
                <td  colspan ="3">
                    <asp:TextBox ID="usr_remarks" runat="server" TextMode="MultiLine" Width="400" Height="80" />
                </td>
            </tr>
            <tr>
                <td colspan="4" align="left" width="100%">
                    <table border="1" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                    <asp:Label ID="lbl_sys_cb" runat="server" />:
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                    <asp:Label ID="sys_cb" runat="server" />
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                    <asp:Label ID="lbl_sys_lub" runat="server" />:
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                    <asp:Label ID="sys_lub" runat="server" />
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                    <asp:Label ID="lbl_sys_cd" runat="server" />:
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                    <asp:Label ID="sys_cd" runat="server" />
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                    <asp:Label ID="lbl_sys_lud" runat="server" />:
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                    <asp:Label ID="sys_lud" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../cms_search.aspx?menu_code=<%=fun_code%>'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div id="div2">
    </div>
    <asp:HiddenField ID="user_type" runat="server" />
    </form>
</body>
</html>
