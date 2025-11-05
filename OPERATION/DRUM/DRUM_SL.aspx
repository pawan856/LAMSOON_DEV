<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DRUM_SL.aspx.vb" Inherits="OPERATION_DRUM_DRUM_SL" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Short-Length Cable</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
      <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 100%">
           <tr>
                <td class="TITLE" colspan="4" align="left">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE" >
                                <b><asp:Label ID="lbl_ImageHd" runat="server" Text="Short-length Cable Details:" /></b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                 <td class="LabelTD" nowrap width="100px">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_ID" runat="server" Text="Drum ID"  /></font>
                </td>
                <td colspan="3">
                    <asp:label runat="server" ID="DRUM_ID" />
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Short-length Cable Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="ILBS_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:BoundField runat="server" DataField="SYS_LUD" HeaderText="Last Update Date" DataFormatString="{0:dd-MM-yyyy hh:mm:ss}" />
                            <asp:BoundField runat="server" DataField="ITM_SKU_NO" HeaderText="Stock No." />
                            <asp:BoundField runat="server" DataField="pack_key" HeaderText="Pack Key" />
                            <asp:BoundField runat="server" DataField="ILBS_SERIAL_NO" HeaderText="Serial No" />
                            <asp:BoundField runat="server" DataField="ILOC_WH" HeaderText="Warehouse" />
                            <asp:BoundField runat="server" DataField="ILOC_LOC" HeaderText="Location" />
                            <asp:BoundField runat="server" DataField="ILBS_DRUM_LEVEL" HeaderText="Drum LV" />
                            <asp:BoundField runat="server" DataField="ILBS_QTY2" HeaderText="QTY2" ItemStyle-HorizontalAlign="Right" />
                            <asp:BoundField runat="server" DataField="ILBS_UOM2" HeaderText="UOM2" />
                            <asp:BoundField runat="server" DataField="ILBS_KG" HeaderText="QTY3" />
                            <asp:BoundField runat="server" DataField="ILBS_UOM3" HeaderText="UOM3" />
                        </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD" align="right">
                    <input id="btnClose" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                        <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="javascript:window.close();"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
