<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UserGPMain.aspx.vb" Inherits="UserGPMain" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>User Group Maintenance</title>
<script language="javascript">
            function DisableDeleteButton() {
             var grp = document.getElementsByName("btnDelete");
             var count;
             count = grp.length;
              if (count == 1) {
                   grp[0].disabled = true;
              }
            }
</script>

<script language ="javascript" src="../js/validation.js"></script>
<script language ="javascript" src="../js/JS_Calendar.js"></script>

<link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css">
<link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
<link rel="stylesheet" href="../stylesheet/general.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="javascript:DisableDeleteButton();">
    <br />
    <form method="post" name="myform" id="myform" onsubmit="" runat="server">
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
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
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='UserGPSearchFun.aspx'"
                        class="all_button" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_grp_code" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="grp_code" runat="server"></asp:TextBox><% If Session("pagemode") = "N" Then%><font color="red">*</font><% End If%>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_grp_name" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="grp_name" runat="server"></asp:TextBox><font color="red">*</font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_grp_type" runat="server" />
                    </font>
                </td>
                <td>
                    <asp:DropDownList ID="grp_type" runat="server">
                        <asp:ListItem Text="User Group" Value ="USER_GROUP"></asp:ListItem>
                    </asp:DropDownList>
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
                <td class="SUBTITLE" colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="SUBTITLE">
                                <b>
                                    <asp:Label ID="userRoleHeader" runat="server" /></b>
                                    &nbsp;
                                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />
                            </td>
                            <td width="100%" class="SUBTITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="usr_id" HorizontalAlign="Left">
                        <Columns>
                            <asp:TemplateField HeaderText="User ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID ="usr_id" runat="server"></asp:DropDownList>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Remove" CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>
                                <ControlStyle Width="60px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
             <tr>
                <td class="SUBTITLE" colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="SUBTITLE">
                                <b>
                                    <asp:Label ID="lblPrgm" runat="server" /></b>
                                    &nbsp;
                                    <asp:Button ID="addPrgm" runat="server" Text="Add" CssClass="all_button" />
                            </td>
                            <td width="100%" class="SUBTITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
             <tr>
                <td colspan="6">
                    <asp:GridView ID="GridView2" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="fun_code" HorizontalAlign="Left">
                        <Columns>
                             <asp:TemplateField HeaderText="Program ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID ="fun_code" runat="server"></asp:DropDownList>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Read" Visible = "false">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="sec_read" runat="server" Checked="false" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Write">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="sec_write" runat="server" Checked="false" />                                   
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Report">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="sec_report" runat="server" Checked="false" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Print">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="sec_print" runat="server" Checked="false" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Access">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="sec_access" runat="server" Checked="false" />                                   
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete_P" name="btnDelete_P" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='UserGPSearchFun.aspx'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
