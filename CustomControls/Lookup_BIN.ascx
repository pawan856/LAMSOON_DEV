<%@ Control Language="VB" AutoEventWireup="false" CodeFile="Lookup_BIN.ascx.vb" Inherits="CustomControls_WebUserControl" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<asp:UpdatePanel runat="server" ID="TEXTBOXUDP" RenderMode="Inline">
    <ContentTemplate>
        <asp:label ID="dsp_LOCATION_CODE" runat="server"  style="display:inline-block;  min-width: 100px;" /><asp:ImageButton ID="btnSearch" runat="server" ImageUrl="~/images/btn_search.gif" ALIGN="center"  />
        <asp:HiddenField runat="server" ID="LOCATION_CODE" />
        <asp:Panel ID="pnlDropdown" runat="server" CssClass="modalPopup" Style="display: none">
                    <table width="640px">
                        <tr>
                            <td class="LabelTD" nowrap width="15%" align="right">Warehouse:</td>
                            <td><asp:DropDownList runat="server" ID="WH_CODE" AutoPostBack="true" /></td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap width="15%" align="right">Floor:</td>
                            <td><asp:DropDownList runat="server" ID="FL_NUM" AutoPostBack="true" />
                            </td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap width="15%" align="right">Area:</td>
                            <td><asp:DropDownList runat="server" ID="AR_CODE" AutoPostBack="true" />
                            </td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap width="15%" align="right">Rack:</td>
                            <td><asp:DropDownList runat="server" ID="RK_CODE" AutoPostBack="true" />
                            </td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap width="15%" align="right">Bin:</td>
                            <td><asp:DropDownList runat="server" ID="BN_CODE" AutoPostBack="true" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                            <asp:Button runat="server" ID="btnOK" CssClass="all_button" Text="OK" />&nbsp;&nbsp;
                            <asp:Button runat="server" ID="btnCancel" CssClass="all_button" Text="Cancel" />
                            </td>
                        </tr>
                    </table>
        </asp:Panel>
        <asp:ModalPopupExtender ID="BIN_ModalPopupExtender" 
            runat="server" 
            DynamicServicePath=""
            Enabled="True" 
            TargetControlID="dummy" 
            PopupControlID="pnlDropdown" 
            BackgroundCssClass="modalBackground"
            DropShadow="true" 
            CancelControlID="btnCancel"
            RepositionMode="None"
            y="10"
            >
        </asp:ModalPopupExtender>
        <asp:HiddenField ID="dummy" runat="server" />
    </ContentTemplate>
</asp:UpdatePanel>