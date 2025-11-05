<%@ Page Language="VB" AutoEventWireup="false" CodeFile="del_dtl.aspx.vb" Inherits="MASTER_WM_del_dtl" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Delete Warehouse Location</title>
    <script language="javascript" src="../../js/validation.js"></script>
    <script language="javascript" src="../../js/JS_Calendar.js"></script>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="form1" runat="server">
    <div>
    <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
            <tr>
                <td class="TITLE">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lheader" runat="server" text="Warehouse Location Details:"/></b>
                            </td>
                            <td width="100%" class="TITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="menuTD">
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.close();"
                        class="all_button" /></td>
            </tr>    
            <tr>
                <td>
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                  Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                  BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                  CellPadding="3" CaptionAlign="Top" DataKeyNames="wh_code,fl_num,ar_code,rk_code,bn_code" HorizontalAlign="Left">
                                  <RowStyle CssClass="GV" />
                                  <HeaderStyle CssClass="DtlLabel" Font-Bold="False" HorizontalAlign="Left"/>
                                  <Columns>
                                  <asp:BoundField DataField="fl_name" HeaderText="Floor" ItemStyle-HorizontalAlign="Left" />
                                  <asp:BoundField DataField="ar_name" HeaderText="Area" ItemStyle-HorizontalAlign="Left" />
                                  <asp:BoundField DataField="rk_name" HeaderText="Rack" ItemStyle-HorizontalAlign="Left" />
                                  <asp:BoundField DataField="bn_code" HeaderText="Bin" ItemStyle-HorizontalAlign="Left" />
                                  <asp:TemplateField HeaderStyle-Width="100px">
                                    <ItemTemplate>
                                          <asp:Button runat="server" ID="btnDelDtl" Text="Delete" CssClass="all_button" CommandName="DEL" />
                                          <asp:HiddenField runat="server" ID="fl_num" />
                                          <asp:HiddenField runat="server" ID="ar_code" />
                                          <asp:HiddenField runat="server" ID="rk_code" />
                                          <asp:HiddenField runat="server" ID="bn_code" />
                                          <asp:HiddenField runat="server" ID="dtl_type" />
                                    </ItemTemplate>
                                  </asp:TemplateField>
                                  </Columns>
                     </asp:GridView>
                </td>
            </tr>
            <tr>
                <td class="menuTD">
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.close();"
                        class="all_button" />
                </td>
            </tr>
        </table>

    </div>
    </form>
</body>
</html>
