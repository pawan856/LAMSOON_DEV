<%@ Page Language="VB" AutoEventWireup="false" CodeFile="genSerialNo.aspx.vb" Inherits="INBOUND_GR_genSerialNo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Generate Serial No.</title>
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="piform" id="piform" runat="server">
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 500px">
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
                <asp:Button ID="SaveBtn1" runat="server" Text="Save" CssClass="all_button" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <table border="1" cellspacing="0" align="Left" rules="all"  width="100%" style="font-family:Arial;font-size:10px;text-decoration:none;height:10px;width:100%;border-collapse:collapse;">
            <tr>
            <td class="DtlLabel" align="left">
                <asp:Label ID="lbl_grd_itm_code" runat="server" Text="Item Code"></asp:Label></td>
            <td class="DtlLabel" align="left">
                <asp:Label ID="lbl_grd_itm_name" runat="server" Text="Item Name"></asp:Label></td>
            <td class="DtlLabel" align="left">
                <asp:Label ID="lbl_sn_prefix" runat="server" Text="Serial No. Prefix" Width="150px"></asp:Label></td>
            <td class="DtlLabel" align="left">
                <asp:Label ID="lbl_sn_digit" runat="server" Text="Serial No. Start From"
                    Width="94px"></asp:Label></td>
            <td class="DtlLabel" align="Right">
                <asp:Label ID="lbl_grd_rcv_qty" runat="server" Text="Received Qty" Width="80px"></asp:Label></td>
            <td class="DtlLabel" align="left">&nbsp;</td>
            </tr>
            <tr>
            <td align="left">
                <asp:Label ID="grd_itm_code" runat="server" Font-Size="11px" Width="130px"></asp:Label>
            </td>
            <td align="left">
                <asp:Label ID="grd_itm_name" runat="server" Font-Size="11px" Width="210px"></asp:Label>
            </td>
            <td align="left">
                <asp:Textbox ID="sn_prefix" runat="server" Font-Size="11px" Width="130px"></asp:Textbox>
            </td>
            <td align="left" nowrap>
                <asp:Textbox ID="sn_digit" runat="server" Font-Size="11px" Width="50px"></asp:Textbox>
                <asp:Label ID="sn_digit_to" runat="server" Font-Size="10px" Text="(1>999)" />
            </td>
            <td align="Right">
                <asp:Label ID="grd_rcv_qty" runat="server" Font-Size="11px" style="text-align:right"></asp:Label>
            </td>
            <td align="left">
                <asp:Button ID="serial_no" text="Gen Serial No." runat="server" Font-Size="11px" CssClass="all_button"></asp:Button>
            </td>
            </tr>
            </table>
        </td>
    </tr>
     <tr runat="server" id="snHtr" visible="false">
        <td class="TITLE" colspan="8">
            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="100%" class="TITLE">
                        <b><asp:Label ID="lbl_ImageHd" runat="server" /></b>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
     <tr runat="server" id="snStr" visible="false">
        <td colspan="8" class="menuTD">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                &nbsp;</td>
        </tr>
        </table>
        </td>
    </tr>
    <tr runat="server" id="snGVtr" visible="false">
        <td colspan="8">
            <asp:GridView ID="SN_GV" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left">
                <Columns>
                    <asp:TemplateField HeaderText="Seq.">
                        <ControlStyle Width="20px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="GRS_SEQ" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="GRS_ITM_CODE" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Serial No.">
                        <ControlStyle Width="130px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="GRS_SERIAL_NO" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Date">
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td>
                                        <asp:Textbox ID="GRS_DATE_IN" runat="server" Font-Size="11px" Width="80" MaxLength="10"></asp:Textbox>
                                    </td>
                                    <td>
                                         <a id="l_GRS_DATE_IN" runat="server" target="_self" style="cursor:hand">
                                            <img id="Img_GRS_DATE_IN" src="../../images/calendar.gif" name="Image1" border=0 align="absmiddle" runat="server">
                                        </a>
                                    </td>
                                </tr>
                            </table>    
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Warranty Date">
                        <ItemTemplate>
                            <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td>
                                        <asp:Textbox ID="GRS_WARR_STDATE" runat="server" Font-Size="11px" Width="80" MaxLength="10"></asp:Textbox>
                                    </td>
                                    <td>
                                         <a id="l_GRS_WARR_STDATE" runat="server" target="_self" style="cursor:hand">
                                            <img id="Img_GRS_WARR_STDATE" src="../../images/calendar.gif" name="Image1" border=0 align="absmiddle" runat="server">
                                        </a>
                                    </td>
                                </tr>
                            </table>    
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Warranty Type">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="GRS_WARR_TYPE" runat="server" Font-Size="11px" maxlength="20"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Warranty Period">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="GRS_WARR_PERIOD" runat="server" Font-Size="11px" maxlength="20"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Warranty Expire Date">
                        <ItemTemplate>
                             <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td>
                                        <asp:Textbox ID="GRS_WARR_EXPDATE" runat="server" Font-Size="11px" Width="80" MaxLength="10"></asp:Textbox>
                                    </td>
                                    <td>
                                         <a id="L_GRS_WARR_EXPDATE" runat="server" target="_self" style="cursor:hand">
                                            <img id="Img_GRS_WARR_EXPDATE" src="../../images/calendar.gif" name="Image1" border=0 align="absmiddle" runat="server">
                                        </a>
                                    </td>
                                </tr>
                            </table>               
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Customer Name">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="CUST_NAME" runat="server" Font-Size="11px" maxlength="100"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Consumer No">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="CONS_NO" runat="server" Font-Size="11px" maxlength="20"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Consumer Name">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="CONS_NAME" runat="server" Font-Size="11px" maxlength="100"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Consumer Tel.">
                        <ControlStyle Width="60px" />
                        <ItemTemplate>
                            <asp:Textbox ID="CONS_TEL" runat="server" Font-Size="11px" maxlength="20"></asp:Textbox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle CssClass="DtlLabel" />
            </asp:GridView>
        </td>
    </tr>
     <tr runat="server" id="snSavetr" visible="false">
        <td colspan="8" class="menuTD">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                <asp:Button ID="savebtn2" runat="server" Text="Save" CssClass="all_button" />
                <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="關閉" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
</table>

<asp:HiddenField ID="editMode" runat="server" />
<asp:HiddenField ID="moduleAction" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
