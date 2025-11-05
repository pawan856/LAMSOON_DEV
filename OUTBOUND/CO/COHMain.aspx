<%@ Page Language="VB" AutoEventWireup="false" CodeFile="COHMain.aspx.vb" Inherits="OUTBOUND_CO_HOLD" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Hold Stock</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">

    function LocLookUp(lb_id, hd_id) {
        if (document.myform.editMode.value != "V") {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.RT_WH.value);
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
            document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
            document.hiddenForm.target = "locLookUp";
            document.hiddenForm.submit();
        }
    }

    function ROLookUp(ITM_STR, RO_ID, SEQ_ID, QTY_ID) {
        if (document.myform.editMode.value != "V") {

            var ori_ro;

            ori_ro = "(" + document.getElementById(RO_ID).value + "||" + document.getElementById(SEQ_ID).value + "),";

            removeAllElementFromForm(document.hiddenForm);

            window.open("", "ROLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=800,left=20,top=35");

            setInterfaceDataToForm(document.hiddenForm, "ITM_STR", ITM_STR);
            setInterfaceDataToForm(document.hiddenForm, "RO_ID", RO_ID);
            setInterfaceDataToForm(document.hiddenForm, "SEQ_ID", SEQ_ID);
            setInterfaceDataToForm(document.hiddenForm, "QTY_ID", QTY_ID);
            setInterfaceDataToForm(document.hiddenForm, "ori_RO", ori_ro);
            //setInterfaceDataToForm(document.hiddenForm, "ori_RO", "(" & document.getElementById(RO_ID).value) & "||" & document.getElementById(SEQ_ID).value & "),";
            document.hiddenForm.action = "ROLookup.aspx";
            document.hiddenForm.target = "ROLookUp";
            document.hiddenForm.submit();
        }
    }

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform"  runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 100%">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                    <asp:label ID="lheader" runat ="server" />
                 </td>
            </tr>            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="btnSave1" runat="server" Text="Save" CssClass="all_button" />
                        <input type="button" id="btnClose" value="Close" class="all_button"  onclick="javascript:window.close();"/>
                       </td>   
                </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CO_code" runat="server" />:</font>
                </td>
                <td colspan="3">
                    <font size="2">
                         <asp:label ID="co_code" runat="server" />
                    </font>
                </td>               
             </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_cod_itm_code" runat="server" />:</font></td>
                <td runat="server">
                    <font size="2">
                        <asp:label ID="cod_itm_code" runat="server" />
                        <asp:HiddenField runat="server" ID="COD_PACK_KEY" />
                        <asp:HiddenField runat="server" ID="COD_PALLET_NO" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                <font size="2">
                    <asp:Label ID="lbl_cod_itm_desc" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:label ID="cod_itm_desc" runat="server"  />
                </td>
              </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_cod_batch_no" runat="server" />:</font></td>
                <td id="Td1" runat="server" colspan="3">
                    <font size="2">
                        <asp:label ID="cod_batch_no" runat="server" />
                    </font>
                </td>         
              </tr>
             <tr style="display:none">
                <td class="LabelTD" nowrap width="15%">
                <font size="2">
                    <asp:Label ID="lbl_Bal_QTY" runat="server"></asp:Label>:</font>
                </td>
                <td>
                    <asp:label ID="Bal_QTY" runat="server"  />
                </td>    
             </tr>
             <tr>
                <td colspan="4">
                    &nbsp;
                 </td>
             </tr>  
             <tr>
                <td colspan="4" class="TITLE">
                    <asp:label ID="lbl_avail_stock" Text="Available Stock" runat ="server" />
                 </td>
            </tr>     
            <tr>
                <td colspan="4">
                    <asp:GridView ID="gvAvailStock" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" ShowHeaderWhenEmpty="true" 
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="batch_no" HorizontalAlign="Left">
                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    <RowStyle CssClass="GV" />
                    <Columns>
                        <asp:boundfield datafield="batch_no" headertext="Batch No." HtmlEncode="false" 
                                        HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false" />
                        <asp:boundfield datafield="avail_qty" headertext="Available Qty" HtmlEncode="false" 
                                        HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false" />
                    </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                &nbsp;
                </td>
            </tr>  
             <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button runat="server" ID="btnAdd" CssClass="all_button" Text="Add Hold Item" />
                 </td>
            </tr>  
            <tr>
                <td colspan="4">
                   <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" ShowHeaderWhenEmpty="true" 
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="COH_SEQ" HorizontalAlign="Left">
                   <RowStyle CssClass="GV" />
                    <Columns>
                            <asp:TemplateField HeaderText="Batcho No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="COH_BATCH_NO" AutoPostBack="true" OnSelectedIndexChanged="ddl_TypeChanged" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Type">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="COH_TYPE" AutoPostBack="true" OnSelectedIndexChanged="ddl_TypeChanged" Font-Size="11px">
                                        <asp:ListItem Text="Hold Stock" Value="STOCK" />
                                        <asp:ListItem Text="Hold RO" Value="RO" />
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Status">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:HiddenField runat="server" ID="ori_status" />
                                    <asp:DropDownList runat="server" ID="COH_STATUS" Font-Size="11px">
                                        <asp:ListItem Text="Hold" Value="HOLD" />
                                        <asp:ListItem Text="Release" Value="RELEASE" />
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Hold Date">
                                <HeaderStyle HorizontalAlign="Left" Width="10%" Wrap="false" />
                                <ItemTemplate>
                                   <asp:Label runat="server" ID="SYS_CD"  Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Release Date">
                                <HeaderStyle HorizontalAlign="Left" Width="10%" Wrap="false" />
                                <ItemTemplate>
                                   <asp:Label runat="server" ID="COH_REL_DATE"  Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="RO No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                   <asp:textbox runat="server" ID="RO_CODE" Font-Size="11px" MaxLength="20" />
                                   <asp:HiddenField runat="server" ID="ROD_SEQ" />
                                   <asp:Image ID="Image_RO_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />   
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Avaliable RO Qty.">
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                   <asp:textbox runat="server" ID="ROD_QTY"  Font-Size="11px"  style="text-align:right" BorderStyle="None" BackColor="#F0F0F0" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Hold Qty.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                   <asp:textbox runat="server" ID="COH_QTY"  Font-Size="11px" MaxLength="15" CssClass="REQUIRED" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="In-stock Qty.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                   <asp:Label runat="server" ID="COH_IN_STOCK_QTY" style="text-align:right" Font-Size="11px"  />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Posted Qty.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                   <asp:Label runat="server" ID="COH_REL_QTY" style="text-align:right" Font-Size="11px"  />
                                </ItemTemplate>
                            </asp:TemplateField>                              
                            <asp:TemplateField>
                                <HeaderStyle HorizontalAlign="Left" Width="5%" />
                                <ItemTemplate>
                                    <asp:Button runat="server" id="btnDEL" Text="Delete" CommandName="DEL" CssClass="all_button" />
                                </ItemTemplate>
                            </asp:TemplateField>
                     </Columns>
                     <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                 </td>
            </tr> 
            <tr>
                <td colspan="4">
                    &nbsp;
                 </td>
            </tr>  
            <tr>
             <td class="menuTD" align="left" colspan="4">
                        <asp:Button ID="btnSave2" runat="server" Text="Save" UseSubmitBehavior="false" CssClass="all_button" />
                        <input type="button" id="btnClose2" value="Close" class="all_button"  onclick="javascript:window.close();"/>
                       </td>   
            </tr>
             <tr>
                <td colspan="4" class="TITLE">
                    &nbsp;
                 </td>
            </tr>  
        </table>
    </div>
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" method="post" />
  </body>
</html>
