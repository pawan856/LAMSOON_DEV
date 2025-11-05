<%@ Page Language="VB" AutoEventWireup="false" CodeFile="wh_spec.aspx.vb" Inherits="MASTER_IM_wh_spec" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script language="javascript">
         function LocLookUp(wh_val, lb_id, hd_id, wh_id) {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");

                setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
                setInterfaceDataToForm(document.hiddenForm, "pForm", "form1");
                if (wh_id)
                    setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_id + "||WH");
                else
                    setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
                document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
                document.hiddenForm.target = "locLookUp";
                document.hiddenForm.submit();        
        }
    </script>
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css">
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
      <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td class="TITLE" colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lheader" runat="server" Text="Warehouse specifications" /></b>
                            </td>
                            <td width="100%" class="TITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD">
                    <asp:Button runat="server" ID="btnAdd" Text="Add" CssClass="all_button" />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                   Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                   BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                   CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                       <Columns>
                          <asp:TemplateField HeaderText="Warehouse">
                              <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="wh_code" CssClass="REQUIRED" />
                              </ItemTemplate>
                              <ItemStyle Font-Size="11px" />
                          </asp:TemplateField>
                          <asp:TemplateField HeaderText="Re-Order Qty.">
                              <ItemTemplate>
                                    <asp:TextBox runat="server" ID="IW_REORD_QTY" width="60px" MaxLength="10" />
                                    <asp:FilteredTextBoxExtender runat="server" ID="filter1" FilterMode="ValidChars" FilterType="Numbers" TargetControlID="IW_REORD_QTY" />
                              </ItemTemplate>
                              <ItemStyle Font-Size="11px" HorizontalAlign="Center" />
                          </asp:TemplateField>
                          <asp:TemplateField HeaderText="Preferred Loc. 1">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td align="left" nowrap>
                                        <asp:Label ID="dsp_IW_PREF_LOC1" width="130" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="IW_PREF_LOC1" runat="server" />
                                    </td>
                                    <td align="right">
                                        <asp:Image ID="Image_Loc_LookUp1" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                                    </td>
                                </tr>
                                </table>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Preferred Loc. 2">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td align="left" nowrap>
                                        <asp:Label ID="dsp_IW_PREF_LOC2" width="130" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="IW_PREF_LOC2" runat="server" />
                                    </td>
                                    <td align="right">
                                        <asp:Image ID="Image_Loc_LookUp2" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                                    </td>
                                </tr>
                                </table>
                                </ItemTemplate>
                            </asp:TemplateField>
                          <asp:TemplateField HeaderText="Picked Loc. ">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                                <tr>
                                    <td align="left" nowrap>
                                        <asp:Label ID="dsp_IW_PICK_LOC" width="130" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="IW_PICK_LOC" runat="server" />
                                    </td>
                                    <td align="right">
                                        <asp:Image ID="Image_Loc_LookUp3" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                                    </td>
                                </tr>
                                </table>
                                </ItemTemplate>
                            </asp:TemplateField>


                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                       </Columns>
                       <HeaderStyle CssClass="TITLE" Font-Size="12px" />
                   </asp:GridView> 
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD" align="right">
                    <asp:Button runat="server" ID="btnSave" Text="Save" CssClass="all_button" />
                    <input type="button" value="Close" onclick="javascript:window.close();" Class="all_button" />
                </td>
            </tr>
       </table>           
    </div>
    </form>
    <form id="hiddenForm" name="hiddenForm" method="post" />
</body>
</html>
