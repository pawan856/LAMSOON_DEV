<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AlertMain.aspx.vb" Inherits="ALERT_AlertMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Alert</title>
    <link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css" />

    <script language ="javascript" src="../js/validation.js"></script>
    <script language ="javascript" src="../js/JS_Calendar.js"></script>
    <script language="javascript" src="../js/listUtil.js"></script>
    <script language="javascript" src="../js/formatUtil.js"></script>
    <script language="javascript" src="../js/formPostInterfacing.js"></script>
    <script src="../../js/jquery-1.9.1.min.js"></script>
    <script src="../../js/jquery-migrate-1.1.1.min.js"></script> 
    <div runat="server" id="DIVSCRIPT">
    <script language="javascript">
        function goLocation(addr) {
            opener.window.location.href = addr;
        }
    </script>
    </div>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
         <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
            <tr>
                <td class="TITLE">
                    <asp:Label runat="server" ID="lHeader" />                    
                </td>
            </tr>           
            <tr>
                <td>
                    <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
                        <tr>
                            <td width="15%" class="LabelTD">
                                <asp:Label runat="server" ID="lbl_ALERT_TO" />
                            </td>
                            <td>
                                <asp:DropDownList runat="server" ID="ALERT_TO" />
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="LabelTD">
                                <asp:Label runat="server" ID="lbl_KEY_WORD" />
                            </td>
                            <td>
                                <asp:textbox runat="server" ID="KEY_WORD" Width="350px" />
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="LabelTD">
                                <asp:Label runat="server" ID="lbl_SENT_DATE" />
                            </td>
                            <td>
                            <%If Session("gLang") = "C" Then%> 由 <% Else%> From <%end if %>:
                                <asp:TextBox ID="DATE_FROM" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                                <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../images/calendar1.gif" 
                                       ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                                       TargetControlID="DATE_FROM" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                            <%If Session("gLang") = "C" Then%> 至 <% Else%> From <%end if %>:
                               <asp:TextBox ID="DATE_TO" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                               <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../images/calendar1.gif" 
                                       ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                               <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                                       TargetControlID="DATE_TO" PopupButtonID="btnDATE_ID2" Format="dd/MM/yyyy" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td><asp:Button runat="server" ID="btnSearch" text="Search" CssClass="all_button" /></td>
            </tr>
            <tr>
                <td class="TITLE">                    
                    <div style="float:right">
                    <asp:dropdownlist runat="server" ID="SORT_BY" AutoPostBack="true">
                        <asp:ListItem Value="ALRT_SEND_DATE" Text="Sent Date" />
                        <asp:ListItem Value="ALRT_PRIORITY" Text="Priority" />                        
                    </asp:dropdownlist>
                    <asp:Button runat="server" ID="btnRead" Text="Read and Remove" CssClass="all_button" />                    
                    </div>
                    
                </td>
            </tr>  
            <tr>
                <td>
                    <asp:updatepanel runat="server" ID="MainUDP" RenderMode="Inline">
                    <ContentTemplate>
                     <asp:panel runat="server" id="GVPan" ScrollBars="Auto" style="max-height: 400px">
                     <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Alert Found."
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" ShowHeaderWhenEmpty="true"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="SYS_PK" HorizontalAlign="Left">
                                <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                                <Columns>
                                    <asp:TemplateField ItemStyle-Width="20px" ItemStyle-HorizontalAlign="Center" HeaderText="Read">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemTemplate>
                                            <asp:CheckBox runat="server" Id="CheckYN" />
                                            <asp:HiddenField runat="server" ID="sys_pk" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Priority" ItemStyle-Width="50px">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_PRIORITY" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Subject" ItemStyle-Width="20%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_SUBJECT" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Details" ItemStyle-Width="40%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_MSG" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Alert To">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_TO_USERID" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Alert Sent By">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_SEND_BY" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Storer">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="STO_NAME" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Document No.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:hyperlink ID="SYS_DOC_NO" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Reference No.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_REF_NO" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Alert Sent On">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="ALRT_SEND_DATE" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                      </asp:GridView>                    
                    </asp:panel>
                    </ContentTemplate>
                </asp:updatepanel>
                </td>
            </tr>
            <tr>
                <td class="TITLE" align="right">
                    <asp:Button runat="server" ID="btnClose" OnClientClick="javascipt:window.close();" CssClass="all_button" />
                </td>
            </tr>
         </table>
    </div>
    </form>
</body>
</html>
