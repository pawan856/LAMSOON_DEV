<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PostCheck.aspx.vb" Inherits="OUTBOUND_DO_PostCheck" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Pick List</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">
    function postDO() {

        if (window.opener && !window.opener.closed &&
            window.opener.document.getElementById("JS_MODULE_ID") != null && window.opener.document.getElementById("JS_MODULE_ID").value == "DO_MAIN") {
            window.opener.postOrderOK();
        }
        else {
            alert('The WIT main screen is not found. Please re-open the WIT to post!');
        }
        window.close();
    }

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="form1" name="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 550px">
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
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                <asp:Button ID="btnPost" Text="Post" OnClientClick="return confirm('Are you sure to post this record?');" CssClass="all_button" runat="server" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td class="menuTD" align="right">
                &nbsp;
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:UpdatePanel runat="server" ID="updtPnl_gvPickList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left" >
                <Columns>
                    <asp:TemplateField HeaderText="Seq No.">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dod_disp_seq" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_seq" runat="server" />
                            <asp:HiddenField ID="pld_item_no" runat="server" />
                            <asp:HiddenField ID="pld_pack_key" runat="server" />
                            <asp:HiddenField ID="pld_pallet_no" runat="server" />
                            <asp:HiddenField ID="pld_batch_no" runat="server" />
                            <asp:HiddenField ID="pld_loc" runat="server" />
                            <asp:HiddenField ID="pld_item_qty" runat="server" />
                            <asp:HiddenField ID="itm_serial_no_yn" runat="server" />
                            <asp:HiddenField ID="itm_type" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>                        
                    <asp:TemplateField HeaderText="Stock No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Serial">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="pld_serial_no" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Serial Check">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_serial_no_chk" CssClass="REQUIRED" runat="server" Font-Size="11px" maxlength="80" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                </Columns>
                <AlternatingRowStyle CssClass="REQUIRED" />
                <RowStyle CssClass="REQUIRED" />
                <EmptyDataRowStyle CssClass="REQUIRED" />
                <EditRowStyle CssClass="REQUIRED" />
            </asp:GridView>
            </ContentTemplate>
            </asp:UpdatePanel>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
            <asp:Button ID="btnPost2" Text="Post" OnClientClick="return confirm('Are you sure to post this record?');" CssClass="all_button" runat="server" />
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

    <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>

    <asp:UpdatePanel runat="server" ID="updtPnlKey">
    <ContentTemplate>
        <asp:HiddenField ID="editMode" runat="server" />
        <asp:HiddenField ID="moduleAction" runat="server" />
    </ContentTemplate>
    </asp:UpdatePanel>

    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="STORER_CODE" runat="server" />
    <asp:HiddenField ID="CO_CODE" runat="server" />
    <asp:HiddenField ID="DO_CODE" runat="server" />

    </form>

    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
