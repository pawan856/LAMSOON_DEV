<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IMPORTEXPORT.aspx.vb" Inherits="IMPORTEXPORT" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title></title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

    <script language="javascript" src="../../js/validation.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/JS_Calendar.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/listUtil.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/formatUtil.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js" type="text/javascript"></script> 
    <script language="javascript" src="//code.jquery.com/jquery-2.1.1.min.js" type="text/javascript"></script>

    <style type="text/css">
        .auto-style1 {
            font-family: "Arial", "Helvetica", "sans-serif";
            font-size: 14px;
            background-color: #d4d0c8;
            color: #193972;
            height: 25px;
        }

        .auto-style2 {
            height: 25px;
        }
              
    </style>

    <script type="text/javascript">
        function ShowProgress() { 
            var dialog = confirm("Are you sure to perform this data operation?");            
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

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
    <br />
    <form id="form1" runat="server">
        <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
        <input type="hidden" name="moduleAction" value="" />

        <div class="loading" align="center" style="display:none">
            Loading. Please wait.<br />
            <br />
            <img src="images/ajax-loader.gif" alt="Loading..." />
        </div>            

        <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black">
            <br />
            <center>
                <table border="0">
                    <tr>
                        <td class="auto-style1" colspan="3" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_IMPORT_EXPORT" runat="server" Font-Bold="True">IMPORT EBS to WMS data</asp:Label>
                            </font>
                        </td>
                        <td class="auto-style2" colspan="3">
                            <font size="2">
                                <asp:DropDownList ID="ddl_IMPORT_EXPORT" runat="server">
                                    <asp:ListItem Text="SELECT" Value="0" Selected />
                                    <asp:ListItem Text="ALL" Value="1" />
                                    <asp:ListItem Text="ITEM MASTER" Value="2" />
                                    <asp:ListItem Text="CUSTOMER MASTER" Value="3" />
                                    <asp:ListItem Text="UOM MASTER" Value="4" />
                                    <asp:ListItem Text="WAREHOUSE MASTER" Value="6" />
                                    <asp:ListItem Text="WAREHOUSE LOCATION" Value="7" />                                    
                                    <asp:ListItem Text="STOCK RETURN" Value="9" />
                                    <asp:ListItem Text="SO" Value="10" />
                                    <asp:ListItem Text="PO" Value="5" />
                                    <asp:ListItem Text="IPR" Value="11" />
                                    <asp:ListItem Text="TO" Value="12" />
                                    <asp:ListItem Text="STOCK ADJ" Value="13" />
                                    <asp:ListItem Text="PICKING RULE" Value="14" />
                                </asp:DropDownList>
                                <asp:Button ID="btnImport" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="EBS IMPORT" OnClick="btnImport_Click" OnClientClick="return ShowProgress()"/>
                             
                            </font>

                        </td>

                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td></td>

                        <td class="auto-style1" colspan="3" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_EXPORT_IMPORT" runat="server" Font-Bold="True">IMPORT WMS to EBS data</asp:Label>
                            </font>
                        </td>

                        <td class="auto-style2" colspan="3">
                            <font size="2">
                                <asp:DropDownList ID="ddl_EXPORT_IMPORT" runat="server">
                                    <asp:ListItem Text="SELECT" Value="0" Selected />
                                    <asp:ListItem Text="ALL" Value="1" />
                                    <asp:ListItem Text="PO GRN" Value="2" />
                                     <asp:ListItem Text="IPR GRN" Value="3" />
                                    <asp:ListItem Text="TO GRN" Value="4" />
                                    <asp:ListItem Text="STOCK RETURN" Value="5" />
                                    <asp:ListItem Text="SO" Value="6" />
                                     <asp:ListItem Text="STOCK RELOCATION" Value="7" />
                                </asp:DropDownList>
                                <asp:Button ID="btnExport" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="WMS Upload" OnClick="btnExport_Click" OnClientClick="return ShowProgress()" />
                               
                            </font>
                        </td>
                    </tr>
                </table>
                <br />
                <asp:Button ID="btnArchiveData" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Archive Data" OnClick="btnArchiveData_Click" OnClientClick="return ShowProgress()"/>
                <asp:Button ID="btnReIndex" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Re-Index Tables" OnClick="btnReIndex_Click" OnClientClick="return ShowProgress()"/>
                <br />

                <p style="margin-left: 2px;">
                    <asp:Literal ID="lblMSG" runat="server"></asp:Literal>
                </p>
            </center>
        </div>

        <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black; display:none;">
            <br />
            <center>
                <table border="0">
                    <tr>
                        <td class="auto-style1" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_STORER_CODE" runat="server" Font-Bold="True">Storer Code</asp:Label>
                            </font>
                        </td>
                        <td class="auto-style2" colspan="3">
                            <font size="2">
                                <asp:DropDownList ID="STORER_CODE" runat="server" MaxLength="20">
                                    <asp:ListItem Text="SELECT" Value="0" Selected />
                                    <asp:ListItem Text="2501-201-LSPS(Flour)" Value="104" />
                                    <asp:ListItem Text="2501-204-LSPS-SF(Flour)" Value="362" />
                                    <asp:ListItem Text="3501-301-LSPS(Oil)" Value="105" />
                                    <asp:ListItem Text="4503-401-LSCCL(Home Care)" Value="106" />
                                </asp:DropDownList>
                            </font>
                        </td>
                    </tr>
                </table>
                <br />
                <p>
                    <asp:Button ID="btnGetCC" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Get Daily Cycle Count" OnClick="btnGetCC_Click" OnClientClick="return confirm('Are you sure to Get Daily Cycle Count?')" />
                    &nbsp;&nbsp;
                                   
                 <asp:Button ID="btnGenCC" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Generate Cycle Count Table-2" OnClick="btnGenCC_Click" OnClientClick="return confirm('Are you sure to Generate Cycle Count Table-2?')" />
                </p>

                <p style="margin-left: 2px;">
                    <asp:Literal ID="lblMSGCC" runat="server"></asp:Literal>
                </p>
                <br />
                <p>
                    <asp:Button ID="btnDeletePOGoodsRCV" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Delete PO and Goods Receive data" OnClick="btnDeletePOGoodsRCV_Click" OnClientClick="return confirm('Are you sure to delete PO and Goods Receive data?')" />
                </p>
                <p style="margin-left: 2px;">
                    <asp:Literal ID="lblDeletedPO" runat="server"></asp:Literal>
                </p>

                <hr />
                <p>
                    <table border="0">
                        <tr>
                            <td class="auto-style1" nowrap>
                                <font size="2">
                                    <asp:Label ID="lblSelectTable" runat="server" Font-Bold="True">Select Table</asp:Label>
                                </font>
                            </td>
                            <td class="auto-style2" colspan="3">
                                <font size="2">
                                    <asp:DropDownList ID="ddlDBTable" runat="server" MaxLength="20">
                                    </asp:DropDownList>
                                </font>
                            </td>
                            <td>
                                <asp:Button ID="btnGetData" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Get DB Table Data" OnClick="btnGetData_Click" OnClientClick="return confirm('Are you sure to Load this table data?')" />
                                &nbsp;&nbsp;
                                <asp:Button ID="btnExportGrid" runat="server" Visible="false" BackColor="black" Font-Bold="True" ForeColor="White" Text="Export Table Data" OnClick="btnExportGrid_Click" />

                            </td>

                        </tr>
                    </table>
                    <asp:GridView ID="GridTableData" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="True" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" ShowFooter="true">
                        <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                </p>
            </center>
        </div>

        <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>

    </form>

</body>

</html>
