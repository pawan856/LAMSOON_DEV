<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GRInsp.aspx.vb" Inherits="INBOUND_GR_GRInsp" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Inspection</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">


</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:window.focus();">
<form name="inspForm" id="inspForm" runat="server">
<asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 500px">
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
    <td colspan="2">
        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
            Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
            EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left">
            <RowStyle CssClass="GV" />
            <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
            <Columns>
                <asp:TemplateField HeaderText="Item Code">
                    <ControlStyle Width="60px" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:Label ID="gri_itm_code" runat="server" Font-Size="11px" />
                        <asp:HiddenField ID="grd_seq" runat="server" />
                        <asp:HiddenField ID="gri_seq" runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Stock No.">
                    <ControlStyle Width="100px" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item Name">
                    <ControlStyle Width="170px" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:Label ID="itm_name" runat="server" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Packy Key">
                    <ControlStyle Width="35px" />
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemTemplate>
                        <asp:Label ID="gri_pack_key" runat="server" Font-Size="11px" style="text-align:center" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Pallet No.">
                    <ControlStyle Width="35px" />
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemTemplate>
                        <asp:Label ID="gri_pallet_no" runat="server" Font-Size="11px" style="text-align:center" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Batch No.">
                    <ItemStyle Width="60px" Wrap="false" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:Label ID="gri_batch_no" runat="server" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Insp. Qty">
                    <ControlStyle Width="50px" />
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemTemplate>
                        <asp:Label ID="gri_insp_qty" runat="server" Font-Size="11px" style="text-align:right" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Right" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Pass Qty">
                    <ControlStyle Width="60px" />
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemTemplate>
                        <asp:TextBox ID="gri_pass_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Right" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Rejected Qty">
                    <ControlStyle Width="50px" />
                    <HeaderStyle HorizontalAlign="Right" />
                    <ItemTemplate>
                        <asp:TextBox ID="gri_rej_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Reject Reason">
                    <ControlStyle />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:DropDownList runat="server" ID="gri_rej_reason" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Remarks">
                    <ControlStyle Width="150px" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:TextBox ID="gri_remark" runat="server" Font-Size="11px" maxlength="4000" Width="150px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Inspected By">
                    <ItemStyle Width="60px" Wrap="false" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:DropDownList runat="server" ID="gri_inspected_by" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Status">
                    <ItemStyle Width="60px" Wrap="false" />
                    <HeaderStyle HorizontalAlign="Left" />
                    <ItemTemplate>
                        <asp:DropDownList runat="server" ID="gri_status" Font-Size="11px" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <AlternatingRowStyle CssClass="REQUIRED" />
            <RowStyle CssClass="REQUIRED" />
            <EmptyDataRowStyle CssClass="REQUIRED" />
            <EditRowStyle CssClass="REQUIRED" />
        </asp:GridView>
    </td>
</tr>
<tr>
    <td colspan="2" class="menuTD" style="height:25px">
        <asp:Button ID="btnSave" Text="Save" CssClass="all_button" runat="server" />
        <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
            <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
            class="all_button" />
    </td>
</tr>

</table>


    <!-- Update panel for asyn postback -->
    <asp:UpdatePanel runat="server" ID="updtPnlPostBack" UpdateMode="Conditional">
    <ContentTemplate />
    </asp:UpdatePanel>

    <!-- Update panel that decided for validation alert -->
    <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>


    <!-- Update panel that decided for key fields and moduleAction -->
    <asp:UpdatePanel runat="server" ID="updtPnlKey">
    <ContentTemplate>
        <asp:HiddenField ID="IMP_CODE" runat="server" />
        <asp:HiddenField ID="STORER_CODE" runat="server" />
        <asp:HiddenField ID="GR_CODE" runat="server" />
        <asp:HiddenField ID="GR_STATUS" runat="server" />
        <asp:HiddenField ID="editMode" runat="server" />
        <asp:HiddenField ID="moduleAction" runat="server" />
    </ContentTemplate>
    </asp:UpdatePanel>
</form>
</body>
</html>
