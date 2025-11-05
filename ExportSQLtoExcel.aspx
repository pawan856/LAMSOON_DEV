<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ExportSQLtoExcel.aspx.vb" Inherits="ExportSQLtoExcel" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>


<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
  <%--  <a href="Default.aspx">Default.aspx</a>--%>
    <form id="form1" runat="server">

         <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
        <h2 style="text-align: center">Other Reports (from SQL)</h2>


        <table width="741" border="0" align="center" cellpadding="0" cellspacing="0">
            <asp:HiddenField ID="hdnParameter" runat="server" Value="" />
            <tr>
                <td>
                     <asp:ComboBox ID="ddlSelectForQuery" runat="server" AutoPostBack="true" Font-Size="20px" RenderMode="Inline" AppendDataBoundItems="True"
                                                  MaxLength="300" Width="400" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblDescription" ForeColor="Red" runat="server"></asp:Label>
                    <br />
                    <br />
                </td>
            </tr>
            <tr>

                <td>
                    <input type="hidden" name="moduleAction" value="" />
                </td>
            </tr>
            <tr style="padding: 5px;">
                <td>
                    <textarea id="txtSQL" runat="server" rows="1" cols="180" style="align-content: center"></textarea>
                </td>
            </tr>
            <tr style="padding: 5px;">
                <td>
                    <asp:Literal ID="ltTable" runat = "server" />
                </td>
            </tr>
            <tr>
                <td id="rDropdown" runat="server"></td>
            </tr>
        </table>
        <br />
        <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black; text-align: center">
            <asp:Button ID="btnDownloadExcel" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Download" OnClick="btnDownloadExcel_Click" />
            <asp:Button ID="btnSHowInTable" CausesValidation="False" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Show In table" OnClick="btnSHowInTable_Click" />
            <asp:Button ID="btnReset" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Reset" OnClick="btnReset_Click" />
        </div>
         <label id="lblCount" style="font-size: 20px;color: blue;" runat="server"></label>
        <asp:GridView ID="GridTableData" runat="server" Height="10px" Width="100%" Font-Names="Arial"
            Font-Overline="False" Font-Size="10px" AutoGenerateColumns="True" EmptyDataText="No Record Found."
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
            CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" ShowFooter="true">
            <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
        </asp:GridView>
        <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>
    </form>
</body>
</html>
