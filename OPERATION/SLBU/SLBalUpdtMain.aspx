<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SLBalUpdtMain.aspx.vb" Inherits="OPERATION_SLBU_SLBalUpdtMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Short Length Cable Update Balance</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<div runat="server" id="DIVSCRIPT">
<script language="javascript">
    function LocLookUp(rowIdx, lb_id, hd_id) {
     
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");
     
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
     
    }

    function getLoad() {
        var load_modalPopup = $find('load_ModalPopupExtender');
        load_modalPopup.show();
    }
</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </asp:ToolkitScriptManager>
    <br />
        <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />    
    <asp:HiddenField ID="ListNoList" runat ="server" />   
    <asp:HiddenField ID="seqList" runat="server" /> 
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="2" class="TITLE">                    
                    <asp:label ID="lheader" runat ="server" />
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../user_group_fn.aspx?menu_code=MENU_OP'"                             
                            class="all_button" />                    
                    </td>
                    <td class="menuTD" align="right">
                    <%If Session("pagemode") <> "N" Then%>             
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" />
                    <% End If%>
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>            
             <tr>
                <td class="LabelTD" nowrap width="150px">
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>              
            </tr>           
            <tr>
                <td class="LabelTD" nowrap width="150px">
                    <font size="2">
                        <asp:Label ID="lbl_itm_sku_no" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="itm_sku_no" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>                
            </tr>     
            <tr>
                <td class="LabelTD" nowrap width="150px">
                    <font size="2">
                        <asp:Label ID="lbl_pack_key" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="pack_key" runat="server" MaxLength="30" Width="150px" Text="1"></asp:TextBox>
                    </font>
                </td>                
            </tr>      
            <tr>
                <td colspan="2">
                    <asp:Button ID="btnSearch" runat="server" Text="Search" class="all_button" />
                </td>
            </tr>               
            <tr>
            <td colspan="2">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="TITLE" colspan="2">
                    <asp:Label ID="lbl_ImageHd" runat="server" />
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />
                </td>
            </tr>
            <tr>
                <td colspan="2">                 
                    <asp:updatepanel runat="server" ID="updtGVEdit" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:GridView ID="gvEdit" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top">
                           <RowStyle CssClass="GV" />
                        <Columns>                             
                            <asp:TemplateField HeaderText="Stock No.">
                                <ControlStyle Width="80px" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" Width="80px"></asp:Label>
                                    <asp:hiddenfield ID="ITM_CODE" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                                                            
                            <asp:TemplateField HeaderText="Pack Key">
                                <ItemTemplate>
                                    <asp:Label ID="PACK_KEY" runat="server" Font-Size="11px" Width="50px"></asp:Label>                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ItemTemplate>
                                    <asp:Label ID="itm_name" runat="server" Font-Size="11px" Width="200px"></asp:Label>                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle />
                                <ItemTemplate>
                                    <asp:Textbox ID="ILOC_BATCH_NO" runat="server" Font-Size="11px" Width="100px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                                                  
                            <asp:TemplateField HeaderText="Serial No.">
                              <ItemTemplate>
                                  <asp:TextBox runat="server" id="ILBS_SERIAL_NO" Width="100px" CssClass="REQUIRED" Font-Size="11px" />                                  
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Center" />
                              <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>       
                            <asp:TemplateField HeaderText="Qty2">
                                <ItemTemplate>
                                    <asp:TextBox ID="ILBS_QTY2" runat="server" Width="60px" CssClass="REQUIRED" Font-Size="11px" MaxLength="15"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                            
                            <asp:TemplateField HeaderText="UOM2">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="ILBS_UOM2" Font-Size="11px" Width="50px" />                                  
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Center" />
                              <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                  
                            <asp:TemplateField HeaderText="Location">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_ILOC_LOC" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td>
                                                    <asp:Image ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:pointer" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="ILOC_LOC" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>                                                                                       
                            <asp:TemplateField HeaderText="Drum ID">
                              <ItemTemplate>
                                  <asp:TextBox runat="server" id="ILBS_DRUM_ID" Font-Size="11px" Width="80px" />
                              </ItemTemplate>                              
                              <ItemStyle HorizontalAlign="Center" />
                              <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Level">
                              <ItemTemplate>
                                  <asp:TextBox runat="server" id="ILBS_DRUM_LEVEL" Font-Size="11px" Width="40px" />                   
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Center" />
                              <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                                                                                                                   
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                            </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                    </ContentTemplate>
                    </asp:updatepanel>                    
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="TITLE" colspan="2">
                    <asp:Label ID="lbl_GV_Header" runat="server" />
                </td>
            </tr>
            <tr>
                <td colspan="2">                 
                    <asp:updatepanel runat="server" ID="updtGVView" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:GridView ID="gvView" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top">
                           <RowStyle CssClass="GV" />
                        <Columns>                             
                            <asp:TemplateField HeaderText="Stock No.">
                                <ControlStyle Width="75px" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" Width="75px"></asp:Label>
                                    <asp:hiddenfield ID="ITM_CODE" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                                                            
                            <asp:TemplateField HeaderText="Pack Key">
                                <ItemTemplate>
                                    <asp:Label ID="PACK_KEY" runat="server" Font-Size="11px" Width="50px"></asp:Label>                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ItemTemplate>
                                    <asp:Label ID="itm_name" runat="server" Font-Size="11px" Width="195px"></asp:Label>                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle />
                                <ItemTemplate>
                                    <asp:Label ID="ILOC_BATCH_NO" runat="server" Font-Size="11px" Width="100px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                                                  
                            <asp:TemplateField HeaderText="Serial No.">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="ILBS_SERIAL_NO" Width="100px" Font-Size="11px"  />                                  
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>       
                            <asp:TemplateField HeaderText="Qty2">
                                <ItemTemplate>
                                    <asp:Label ID="ILBS_QTY2" runat="server" Width="60px" Font-Size="11px" MaxLength="15" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                            
                            <asp:TemplateField HeaderText="UOM2">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="ILBS_UOM2" Font-Size="11px" Width="50px" />                                  
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Center" />
                                <HeaderStyle HorizontalAlign="Center" />
                            </asp:TemplateField>                  
                            <asp:TemplateField HeaderText="Location">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                   <asp:Label ID="ILOC_LOC" width="125px" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>                                                                                       
                            <asp:TemplateField HeaderText="Drum ID">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="ILBS_DRUM_ID" Font-Size="11px" Width="80px" />
                              </ItemTemplate>                              
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Level">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="ILBS_DRUM_LEVEL" Font-Size="11px" Width="95px" />                   
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                                                                                                                                               
                            </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                    </ContentTemplate>
                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="newrow" EventName="Click" />
                        <asp:AsyncPostBackTrigger ControlID="btnPost" EventName="Click" />
                    </Triggers>
                    </asp:updatepanel>                    
                </td>
            </tr>
            <tr>
                <td colspan="2" class="menuTD">
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../user_group_fn.aspx?menu_code=MENU_OP'"
                        class="all_button" />
                </td>
            </tr>
        </table>
        <asp:HiddenField runat="server" ID="dummy" />
    </div>

<asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
                ScrollBars="None">
                <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
                    height: 80px">
                    <tr>
                        <td align="center" style="background-color: White; width: 100%; height: 80px; vertical-align: middle">                            
                            <font color="#193B65" style="width: 100%; text-align: center; font-size: 16px;">Posting...please wait</font>
                        </td>
                    </tr>                    
                </table>
            </asp:Panel>
            <asp:HiddenField ID="loadDummy" runat="server" />
            <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
                TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
                DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
                Y="250">
            </asp:ModalPopupExtender>

    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>   
</html>
