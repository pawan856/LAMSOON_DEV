<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATTACH_MAIN.aspx.vb" Inherits="TMS_ATTACH" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Upload Attachment</title>
        <meta http-equiv="content-type" content="text/html; charset=UTF-8">
<link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../js/validation.js"></script>
<script language ="javascript" src="../js/JS_Calendar.js"></script>
<script language="javascript" src="../js/listUtil.js"></script>
<script language="javascript" src="../js/formatUtil.js"></script>
<script language="javascript" src="../js/formPostInterfacing.js"></script>

</head>
<body>
    <form id="form1" runat="server">
     <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
    
      <table border="0" cellspacing="1" cellpadding="1" align="center" width="90%">
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
                <td colspan="4"class="menuTD">
                   <asp:Button ID="AddBtn1" runat="server" Text="New File" CssClass="all_button" 
                        UseSubmitBehavior="false" />
                   <asp:Button ID="CloseBtn1" runat="server" Text="Close" CssClass="all_button" OnClientClick="javascript:window.close();return false;"  />
                </td>
            </tr>
            <tr>
           <td colspan="4">
           <asp:GridView ID="GridView1" runat="server" Height="40px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" 
                        AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="UPFL_SYS_FILENAME,UPFL_DOC_NO" 
                        HorizontalAlign="Left" AllowPaging="True">
                        <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:TemplateField HeaderText="Create Date">
                                <ItemTemplate>
                                    <asp:label ID="SYS_CD" runat="server" />
                                </ItemTemplate>
                            <ItemStyle HorizontalAlign="center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Document Title">
                                <ItemTemplate>
                                    <asp:label ID="UPFL_TITLE" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="File Name">
                                <ItemTemplate>
                                    <asp:LinkButton ID="UPFL_USER_FILENAME" runat="server" CommandName="EDITInfo" />
                                    <asp:HiddenField ID="UPFL_SYS_FILENAME" runat="server"  />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Document Size">
                            <ItemTemplate>
                                    <asp:label ID="UPFL_FILESIZE" runat="server" />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description">
                            <ItemTemplate>
                                    <asp:label ID="UPFL_DESC" runat="server" />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="left" />
                            </asp:TemplateField>
                           <asp:TemplateField>
                            <ItemTemplate>
                                    <asp:button ID="btnDL" runat="server" Text="Download" CssClass="all_button" CommandName="DLFile" />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="center" />
                           </asp:TemplateField>
                            <asp:TemplateField>
                               <ItemTemplate>
                                  <asp:button ID="btnDEL" runat="server" Text="Delete" CssClass="all_button" CommandName="DelFile" OnClientClick="if(!confirm('Are you sure to Delete this File?')) return false;"/>
                               </ItemTemplate>
                            <ItemStyle HorizontalAlign="center" />
                           </asp:TemplateField>
                        </Columns>
                        <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                        <PagerTemplate>
                <table cellspacing="0" cellpadding="0" border="0" width="100%">
                <tr>
                <td align="center">
                    <asp:ImageButton ID="pager_first" ImageUrl="../images/arrow_frist.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="First" />
                    <asp:ImageButton ID="pager_previous" ImageUrl="../images/arrow_previous.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Prev" />
                    <asp:Label ID="lblPager" Text="Page : " Font-Bold="true" style="vertical-align:middle" runat="server" />
                    <asp:DropDownList ID="pager_select" AutoPostBack="true" Width="50" style="vertical-align:middle" runat="server" OnSelectedIndexChanged="GridView1_PageIndexChanged" />
                    <asp:ImageButton ID="pager_next" ImageUrl="../images/arrow_next.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Next" />
                    <asp:ImageButton ID="pager_last" ImageUrl="../images/arrow_last.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Last" />                   
                </td>
                </tr>
                </table>
                </PagerTemplate>
           </asp:GridView>
           </td>
           </tr>
           </table>
           
   <asp:Panel ID="pnlAmend" runat="server" CssClass="modalPopup" Style="display: none">
    <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
        <tr>
            <td class="DtlLabel" height="25" style="background-color:#92A1B9;text-align:center;">
           <asp:Panel runat="Server" ID="PanelDrag" Style="cursor: move;">
                <font style="font-weight:bold; color:Black; font-size: 13px">
                     File Info</font></asp:Panel> 
           </td>
        </tr>
        <tr>
                    <td class="menuTD" align="left">
                     <%--<asp:UpdateProgress runat="server" id="PageUpdateProgress">
                            <ProgressTemplate>
                                Uploading...
                            </ProgressTemplate>
                     </asp:UpdateProgress>
                      <asp:UpdatePanel runat="server" id="udPanel1">
                       <ContentTemplate>
                        
                       </ContentTemplate>
                      </asp:UpdatePanel> --%>
                        <asp:Button ID="pnlSaveBtn" runat="server" Text="Save" CssClass="all_button" OnClientClick="if(!confirm('Are you sure to Save this File info?')) return false;" UseSubmitBehavior="false" />
                        <asp:Button ID="pnlDelBtn" runat="server" Text="Delete" CssClass="all_button"  OnClientClick="if(!confirm('Are you sure to Delete this File?')) return false;" UseSubmitBehavior="false" visible="false"/>
                        <asp:Button ID="btnConfCancel" runat="server" Text="Cancel" CssClass="all_button" />
                    </td>
        </tr>
        <tr>
            <td>
            <table border="0" cellspacing="1" cellpadding="1" align="center"  width="100%">
                <tr> <td class="LabelTD" width="25%"><asp:label id="lbl_Pnl_DOC_TYPE" runat="server" Text="Document Type:" />  </td>
                     <td colspan="3"><asp:label id="DSP_Pnl_DOC_TYPE" runat="server" />
                         <asp:HiddenField runat="server" ID="Pnl_DOC_TYPE" />   
                         <asp:HiddenField runat="server" ID="Pnl_UPFL_SYS_FILENAME" />
                         <asp:HiddenField runat="server" ID="Pnl_UPFL_DOC_NO" />
                         <asp:HiddenField runat="server" ID="EDITMODE" />
                     </td>
                </tr>
                <tr> <td class="LabelTD" width="25%"><asp:label id="lbl_Pnl_FILE_UPLOAD" runat="server" Text="Document Upload:" />  </td>
                     <td>
                     <asp:FileUpload ID="Pnl_FILEUPLOAD" runat="server" />
                         <asp:label id="Pnl_UPFL_USER_FILENAME" runat="server" />
                     </td>
                     <td class="LabelTD" width="25%"><asp:label id="lbl_Pnl_FileSize" runat="server" Text="File Size" /> </td>
                     <td><asp:Label runat="server" ID="Pnl_UPFL_FILESIZE"/> KB</td>
                </tr>
                <tr> <td class="LabelTD" width="25%"><asp:label id="lbl_Pnl_UPFL_TITLE" runat="server" Text="Document Title:" />  </td>
                     <td colspan="3"> 
                     <asp:TextBox ID="Pnl_UPFL_TITLE" runat="server" MaxLength="80" Width="400" />
                     </td>
                    
                </tr>
                <tr>
                <td colspan="4" class="LabelTD"> <asp:Label runat="server" ID="lbl_PNL_UPFL_DESC" Text="Description:" />
                </td>
                </tr>
                <tr><td colspan="4" align="center">
                     <asp:textbox ID="PNL_UPFL_DESC" runat="server" TextMode="MultiLine" Rows="5" MaxLength="500" width="98%"/>
                     </td>
                </tr>
                <tr><td colspan="4">
                    <table border="1" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                <asp:Label ID="lbl_sys_cb" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cb" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lub" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lub" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_cd" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cd" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lud" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lud" runat="server" />
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
                </tr>
            </table>
            </td>
        </tr>
    </table> 
</asp:Panel> 

<asp:HiddenField ID="dummy" runat="server" /> 

<asp:ModalPopupExtender ID="btnAmend_ModalPopupExtender" runat="server" 
    DynamicServicePath="" 
    Enabled="True" 
    TargetControlID="dummy" 
    PopupControlID="pnlAmend"
    BackgroundCssClass="modalBackground"
    DropShadow="true"         
    CancelControlID="btnConfCancel"
    Y="50"
    PopupDragHandleControlID="PanelDrag" >
</asp:ModalPopupExtender>
    </div>
    </form>
</body>
</html>
