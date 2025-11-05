<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RFID_ALERT.aspx.vb" Inherits="OPERATION_RFID_ALERT_RFID_ALERT" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title></title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">


</script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" AsyncPostBackTimeout="3600" />    
    <div>
    <asp:updatepanel runat="server" ID="srcUDP" RenderMode="Inline">
     <ContentTemplate>
     <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td colspan="2" class="TITLE">
                    <asp:label ID="lheader" runat ="server" Text="RBSS Buffer Store Alert Inquiry" />                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="20%" align="right">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_src_Alerts" Text="Status:" />
                    </font>
                </td>
                <td align="left">
                    <asp:CheckBoxList runat="server" ID="src_Alerts" RepeatDirection="Horizontal">
                        <asp:ListItem Value="NEW" Text="New" Selected="True" />
                        <asp:ListItem Value="CHECKED" Text="Checked" />
                    </asp:CheckBoxList>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="20%" align="right">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_src_Status" Text="Alert Type:" />
                    </font>
                </td>
                <td align="left">
                    <asp:CheckBoxList runat="server" ID="src_Status" RepeatDirection="Horizontal">
                        <asp:ListItem Value="STOCK_IN" Text="Stock In" />
                        <asp:ListItem Value="STOCK_OUT" Text="Stock Out" />
                        <asp:ListItem Value="STOCK_RETURN" Text="Stock Return" />
                        <asp:ListItem Value="SURPLUS_IN" Text="Surplus In" />
                        <asp:ListItem Value="OTHERS" Text="OTHERS" />
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="20%" align="right">
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_src_Date" Text="Date:" />
                    </font>
                </td>
                <td align="left">
                    From :
                    <asp:TextBox ID="FR_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="FR_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />&nbsp;&nbsp;
                    To :
                    <asp:TextBox ID="TO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="TO_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
            </tr>
            <tr>
                <td colspan="2">     
                        <asp:button runat="server" ID="btnSearch" CssClass="all_button" Text="Search" />
                        <asp:button runat="server" ID="btnReset"  CssClass="all_button" Text="Reset" />                    
                </td>
            </tr>
     </table>
     </ContentTemplate>        
     <Triggers>
        <asp:PostBackTrigger ControlID="btnReset" />
     </Triggers> 
     </asp:updatepanel>
     <br />     
     <table  border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td  class="TITLE">
                    <asp:updatepanel runat="server" ID="UDPBtn2" RenderMode="Inline">
                    <ContentTemplate>
                        <font size="2">Total of <asp:Label runat="server" ID="total_count" /> records found.</font>
                        &nbsp;&nbsp;
                        <asp:Button runat="server" ID="btnRead" Text="Set New to Checked" CssClass="all_button" />
                    </ContentTemplate>
                </asp:updatepanel>                                        

                <asp:Panel runat="server" style="float:right; background-color:#d4d0c8; padding: 2px;">
                    <asp:Label runat="server" ID="lblRPT" Text="Download as : " ForeColor="Black" Font-Size="11px" />
                        <asp:DropDownList runat="server" ID="reportFormat" Font-Size="11px" >
                            <asp:ListItem Text="Excel" Value="EXCELOPENXML" />
                            <asp:ListItem Text="PDF" Value="PDF" />
                        </asp:DropDownList>
                    <asp:Button runat="server" id="btndownRPT" Text="Download" CssClass="all_button" Font-Size="11px" />
                </asp:Panel>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:updatepanel runat="server" ID="gvUDP" RenderMode="Inline">
                    <ContentTemplate>
                    <asp:GridView ID="Gridview1" runat="server" Height="10px" Width="98%" Font-Names="Arial"  AllowPaging="True" PageSize="25"
                                    Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                    BackColor="White" BorderColor="#000" BorderStyle="Solid" BorderWidth="1px" 
                                    CellPadding="3" CaptionAlign="Top" HorizontalAlign="center">
                    <RowStyle CssClass="GV" />
                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                    <Columns>
                         <asp:BoundField runat="server" DataField="WRMV_BATCH_DATE" HeaderText="Date" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                         <asp:BoundField runat="server" DataField="WRMV_BATCH_TIME" HeaderText="Time" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                         <asp:BoundField runat="server" DataField="WRAL_ITM_SKU_NO" HeaderText="Stock No." HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                         <asp:BoundField runat="server" DataField="WRMV_TAG_ID" HeaderText="RFID" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                         <asp:TemplateField HeaderText="Alert Type">
                                <ItemTemplate>
                                    <asp:Label runat="server" id="WRAL_DSP_ALERT_TYPE" Font-Size="11px" />
                                    <asp:HiddenField runat="server" ID="WRST_CODE" />
                                    <asp:HiddenField runat="server" id="WRAL_KEY" />
                                </ItemTemplate>                                
                         </asp:TemplateField>
                         <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Label runat="server" id="ALERT_DESC" Font-Size="11px" />                                    
                                </ItemTemplate>                                
                         </asp:TemplateField>
                         <asp:TemplateField HeaderText="Video Link" HeaderStyle-Width="80px">
                                <ItemTemplate>
                                    <asp:LinkButton runat="server" ID="video_link" Text="Show Video" CommandName="VIDEO" />                                    
                                </ItemTemplate>                                
                         </asp:TemplateField>
                         <asp:BoundField runat="server" DataField="WRAL_STATUS" HeaderText="Status" HeaderStyle-HorizontalAlign="center" ItemStyle-Font-Size="11px" />
                         <asp:TemplateField HeaderText="Mark as Old" HeaderStyle-Width="50px" >
                                <ItemTemplate>
                                    <asp:checkbox runat="server" ID="chkNew" />
                                </ItemTemplate>                                
                                <ItemStyle HorizontalAlign="Center" />
                         </asp:TemplateField>
                    </Columns>
                    <PagerTemplate>
                           <table cellspacing="0" cellpadding="0" border="0" width="100%">
                            <tr>
                                <td align="center">
                                    <asp:ImageButton ID="pager_first" ImageUrl="../../images/arrow_frist.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="First" />
                                    <asp:ImageButton ID="pager_previous" ImageUrl="../../images/arrow_previous.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Prev" />
                                    <asp:Label ID="lblPager" Text="Page : " Font-Bold="true" style="vertical-align:middle" runat="server" />
                                    <asp:DropDownList ID="pager_select" AutoPostBack="true" Width="50" style="vertical-align:middle" runat="server" OnSelectedIndexChanged="Gridview1_PageIndexChanged" />
                                    <asp:ImageButton ID="pager_next" ImageUrl="../../images/arrow_next.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Next" />
                                    <asp:ImageButton ID="pager_last" ImageUrl="../../images/arrow_last.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Last" />                   
                                </td>
                            </tr>
                        </table>                      
                    </PagerTemplate>                                    
                    </asp:GridView> 
                    </ContentTemplate>
                    </asp:updatepanel>
                </td>
            </tr>
     </table>



    <asp:UpdatePanel runat="server" ID="udpVDO" RenderMode="Inline">
    <ContentTemplate>
         <asp:HiddenField runat="server" ID="dummy" />
    <asp:Panel ID="PnlVDO" runat="server" CssClass="modalPopup" style="display: none"><!-- -->

        <table width="600px">
           <tr>
                <td class="TITLE">
                    <asp:Panel runat="Server" ID="PanelDrag1" Style="cursor: move;">
                    <font size="2">Video List</font>
                    </asp:Panel>
                </td>
           </tr>           
           <tr>
              <td>
              <asp:Panel runat="server" id="gvVPnl" ScrollBars="Auto" style="max-height: 400px; overflow: auto;">
                   <asp:GridView ID="GVVideo" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Video Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                        <Columns>
                            <asp:BoundField DataField="CameraName" HeaderText="Camera" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="videoDate" HeaderText="Video Time" HeaderStyle-Width="35%" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="right" />
                            </asp:BoundField>
                            <asp:TemplateField>
                                 <ItemTemplate>
                                    <asp:LinkButton runat="server" ID="FileName" CommandName="DOWNLOAD" Font-Size="11px" />
                                    <asp:HiddenField runat="server" ID="CameraFolder" />
                                 </ItemTemplate>                                 
                            </asp:TemplateField>                           
                        </Columns>
                   </asp:GridView> 
              </asp:Panel>
              </td>
           </tr>
           <tr>
                <td class="TITLE" align="left">                                        
                    <asp:Button runat="server" ID="btnPnlClose" Text= "Close" CssClass="all_button" />
                </td>
           </tr>
        </table>
    </asp:Panel>
    
    <asp:ModalPopupExtender ID="pnlVDO_ModalPopupExtender" runat="server"
        Enabled="True" 
        TargetControlID="dummy" 
        PopupControlID="PnlVDO"
        BackgroundCssClass="modalBackground"
        DropShadow="true"
        CancelControlID="btnPnlClose"        
        RepositionMode="RepositionOnWindowResizeAndScroll"
        PopupDragHandleControlID="PanelDrag1">
    </asp:ModalPopupExtender> 
    </ContentTemplate>
    <Triggers>
        <asp:PostBackTrigger ControlID="GVVideo" />
    </Triggers>    
    </asp:UpdatePanel>

    </div>
    </form>
</body>
</html>
