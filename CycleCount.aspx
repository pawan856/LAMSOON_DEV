<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CycleCount.aspx.vb" Inherits="CycleCount" EnableEventValidation="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script>

</script>
    <style>
        #tblWavePick, th, td {
            border: 1px solid black;
            width: 50%;
            border-collapse: collapse;
        }

        #tblWavePick {
            width: 60%;
            margin-left: 20%;
            margin-right: 20%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table id="tblWavePick">
                <tr style="background-color: #F0F0F0">
                    <td colspan="2">
                        <h3>CYCLE COUNT </h3>
                    </td>
                </tr>
                <tr>
                    <td>Storer</td>
                    <td>
                        <asp:DropDownList ID="STORER_CODE" runat="server" Width="150px" AutoPostBack="True">
                        </asp:DropDownList>
                        &nbsp;<asp:Button ID="btnCreate" Text="Create" CssClass="all_button" runat="server" OnClientClick="return confirm('Are you sure ?');" />
                    </td>
                </tr>
                <tr>
                    <td>No of record to pick</td>
                    <td>
                        <asp:TextBox ID="txtPickCount" runat="server"></asp:TextBox>
                        &nbsp;<asp:Button ID="btnDailyCount" Text="Daily Cycle Count" CssClass="all_button" runat="server" OnClick="btnDailyCount_Click" OnClientClick="return confirm('Are you sure ?');" />
                    </td>
                </tr>
                <tr>
                    <td style="min-height: 300px">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td>
                        <asp:Label runat="server" ID="lblMsg"></asp:Label></td>
                    <td style="text-align: right">&nbsp;</td>
                </tr>

                <tr>
                    <td>
                        <asp:Button ID="btnListCCItems" Text="List CC Items" CssClass="all_button" runat="server" OnClick="btnListCCItems_Click" />
                    </td>
                    <td style="text-align: right">&nbsp;</td>
                </tr>

                <tr>
                    <td>No of new records</td>
                    <td style="text-align: left">
                        <asp:Label ID="lblNewCount" runat="server"></asp:Label>
                    </td>
                </tr>

                <tr>
                    <td>No of new records having no stock</td>
                    <td style="text-align: left">
                        <asp:Label ID="lblNewCountZeroStk" runat="server"></asp:Label>
                    </td>
                </tr>

                <tr>
                    <td>&nbsp;</td>
                    <td style="text-align: right">
                        <asp:Button ID="btnExport" Text="Export" CssClass="all_button" runat="server" OnClick="btnExport_Click" />
                    </td>
                </tr>

                <tr>
                    <td colspan="2">
                        <asp:GridView ID="gvCCItems" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                            Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                            CellPadding="3" CaptionAlign="Top">
                            <RowStyle CssClass="GV" />
                            <Columns>
                                <asp:BoundField runat="server" DataField="SEQ_NO" HeaderText="SEQ_NO" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="STORER_CODE" HeaderText="STORER_CODE" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="ITEM_CODE" HeaderText="ITEM_CODE" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="ITM_SKU_NO" HeaderText="ITM_SKU_NO" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                                <asp:BoundField runat="server" DataField="ITM_DESC" HeaderText="ITM_DESC" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                                <asp:BoundField runat="server" DataField="ITM_MFG" HeaderText="ITM_MFG" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="ITM_QTY" HeaderText="ITM_QTY" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                                <asp:BoundField runat="server" DataField="WH_CODE" HeaderText="WH_CODE" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="LOT_NO" HeaderText="LOT_NO" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="STATUS" HeaderText="STATUS" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                                <asp:BoundField runat="server" DataField="CDATE" HeaderText="CDATE" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                                <asp:BoundField runat="server" DataField="CByFk" HeaderText="CByFk" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" ItemStyle-Width="100px" HeaderStyle-Width="100px" />
                            </Columns>
                            <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                        </asp:GridView>
                    </td>

                </tr>

            </table>
            <div>
            </div>
            <div>
            </div>

        </div>
    </form>
</body>

</html>
