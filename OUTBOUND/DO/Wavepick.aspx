<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Wavepick.aspx.vb" Inherits="OUTBOUND_DO_Wavepick" EnableEventValidation="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="../../js/formPostInterfacing.js" type="text/javascript"></script> 
    <script src="//code.jquery.com/jquery-2.1.1.min.js" type="text/javascript"></script>

    <script>
        function deleteDo() {
            document.getElementById('hfDODelete').value = 'Y';
            __doPostBack('btnPickingDelete', 'OnClick');
        }

        function ShowProgress() {
            var dialog = confirm("Are you sure ?");
            
            if (dialog == true) {
                setTimeout(function () {
                    var modal = $('<div />');
                    modal.addClass("modal");
                    $('body').append(modal);
                    var loading = $(".loading");
                    loading.show();
                }, 200);
            }
            return dialog;
        }

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
          <div class="loading" style="display:none;text-align:center;">
            Loading. Please wait.<br />
            <br />
            <img src="../../images/ajax-loader.gif" alt="Loading..."/>
        </div>

        <div>
            <table id="tblWavePick">
                <tr style="background-color: #F0F0F0">
                    <td colspan="2">
                        <h3>Customers Orders </h3>
                    </td>
                </tr>
                <tr>
                    <td>Organizations</td>
                    <td>
                        <asp:DropDownList ID="STORER_CODE" runat="server" Width="350px" AutoPostBack="True">
                        </asp:DropDownList>
                        </td>
                </tr>
                <tr>
                    <td>Date</td>
                    <td>Route</td>
                </tr>
                <tr>
                    <td>
                        <asp:Button ID="btnSelectCODate" Text="Select/Deselect All" CssClass="all_button" runat="server" />
                    </td>
                    <td>
                        <asp:Button ID="btnSelectCORoute" Text="Select/Deselect All" CssClass="all_button" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td style="min-height: 300px">
                        <asp:CheckBoxList ID="lboxCODate" runat="server" SelectionMode="Multiple" Width="100%" AutoPostBack="True" BorderStyle="None"></asp:CheckBoxList>
                    </td>
                    <td>
                        <asp:CheckBoxList ID="lboxCORoute" runat="server" SelectionMode="Multiple" Width="100%"></asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td style="text-align: right">
                        <asp:Button ID="btnPicking" Text="Lot Allocation" CssClass="all_button" runat="server"  OnClientClick="return ShowProgress();"/>
                       <%-- OnClientClick="return confirm('Are you sure ?');"--%>
                        <asp:Button ID="btnPickingDeleteCO" Text="Delete CO" CssClass="all_button" runat="server" OnClientClick="return confirm('Are you sure ?');" />
                    </td>
                </tr>

                <tr style="background-color: #F0F0F0">
                    <td colspan="2">
                        <h3>Delivery Orders </h3>
                        <asp:HiddenField ID="hfDODelete" Value="N" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>Date
                    </td>
                    <td>Route
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Button ID="btnSelectDODate" Text="Select/Deselect All" CssClass="all_button" runat="server" />
                    </td>
                    <td>
                        <asp:Button ID="btnSelectRoute" Text="Select/Deselect All" CssClass="all_button" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:CheckBoxList ID="lboxDate" runat="server" SelectionMode="Multiple" Width="100%" AutoPostBack="True"></asp:CheckBoxList>
                    </td>
                    <td>
                        <asp:CheckBoxList ID="lboxRoute" runat="server" SelectionMode="Multiple" Width="100%"></asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Literal ID="ltlMessage" runat="server"></asp:Literal></td>
                    <td style="text-align: right">
                         <asp:Button ID="btnUnAssign" Text="UnAssign" CssClass="all_button" runat="server" OnClientClick="return confirm('Are you sure ?');"/>
                        <asp:Button ID="btnRelease" Text="Release for Picking" CssClass="all_button" runat="server" OnClientClick="return ShowProgress();"/>
                        <%--OnClientClick="return confirm('Are you sure ?');" --%>
                        <asp:Button ID="btnPickingDelete" Text="Delete DO" CssClass="all_button" runat="server" OnClientClick="return confirm('Are you sure ?');" />
                    </td>

                </tr>

            </table>

            <div>
                <asp:Button ID="btnBackUpStock" runat="server" Text="Backup Stock" />
                <asp:Button ID="btnRestoreStock" runat="server" Text="Restore Stock" OnClientClick="return confirm('Are you sure ?');"/>
            </div>

        </div>
    </form>
</body>

</html>
