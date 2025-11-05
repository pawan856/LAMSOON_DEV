<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cc_setup.aspx.vb" Inherits="MASTER_IM_wh_spec" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script language="javascript">
         
    </script>
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css">
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
      <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 600px">
            <tr>
                <td class="TITLE" colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lheader" runat="server" Text="Cycle Count Setup" /></b>
                            </td>
                            <td width="100%" class="TITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="menuTD">
                    <asp:Button runat="server" ID="btnAdd" Text="Add" CssClass="all_button" />
                </td>
                <td class="menuTD" align="right">
                    <asp:Button runat="server" ID="btnSave" Text="Save" CssClass="all_button" />
                    <input type="button" value="Back" onclick="javascript:window.location='../../user_group_fn.aspx?User_ID=&menu_code=MENU_OP';" class="all_button" />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                   Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                   BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" DataKeyNames="CCS_YEAR,CCS_PERIOD"
                   CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                       <Columns>
                          <asp:TemplateField HeaderText="Year" HeaderStyle-HorizontalAlign="Left">
                              <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="CCS_YEAR" CssClass="REQUIRED" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                              </ItemTemplate>
                              <ItemStyle Font-Size="11px" />
                          </asp:TemplateField>
                          <asp:TemplateField HeaderText="Period" HeaderStyle-HorizontalAlign="Left">
                              <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="CCS_PERIOD" CssClass="REQUIRED">
                                        <asp:ListItem Text="1 Year" Value="1" />
                                        <asp:ListItem Text="2 Year" Value="2" />
                                        <asp:ListItem Text="4 Year" Value="4" />
                                    </asp:DropDownList> 
                              </ItemTemplate>
                              <ItemStyle Font-Size="11px" />
                          </asp:TemplateField>
                          <asp:TemplateField HeaderText="End Date" HeaderStyle-HorizontalAlign="Left">
                              <ItemTemplate>
                                    <asp:TextBox ID="CCS_END_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" CssClass="REQUIRED" />
                                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="CCS_END_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />  
                              </ItemTemplate>
                              <ItemStyle Font-Size="11px" />
                          </asp:TemplateField>
                       </Columns>
                       <HeaderStyle CssClass="TITLE" Font-Size="12px" />
                   </asp:GridView> 
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD" align="right">
                    <asp:Button runat="server" ID="btnSave2" Text="Save" CssClass="all_button" />
                    <input type="button" value="Back" onclick="javascript:window.location='../../user_group_fn.aspx?User_ID=&menu_code=MENU_OP';" class="all_button" />
                </td>
            </tr>
       </table>           
    </div>
    </form>
    <form id="hiddenForm" name="hiddenForm" method="post" />
</body>
</html>
