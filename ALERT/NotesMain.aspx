<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NotesMain.aspx.vb" Inherits="ALERT_NotesMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Notes</title>
    <link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css" />

    <script language ="javascript" src="../js/validation.js"></script>
    <script language ="javascript" src="../js/JS_Calendar.js"></script>
    <script language="javascript" src="../js/listUtil.js"></script>
    <script language="javascript" src="../js/formatUtil.js"></script>
    <script language="javascript" src="../js/formPostInterfacing.js"></script>
    <script src="../../js/jquery-1.9.1.min.js"></script>
    <script src="../../js/jquery-migrate-1.1.1.min.js"></script> 
    <div runat="server" id="DIVSCRIPT">
    <script language="javascript">
        function limitText(limitField, limitNum) {
            if (limitField.value.length > limitNum) {
                limitField.value = limitField.value.substring(0, limitNum);
            }
        }
    </script>
    </div>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
            <tr runat="server" visible="false">
                <td align="right">
                    <asp:Label runat="server" ID="usr_id" /> <asp:DropDownList runat="server" ID="TASK_ASSIGN_TO" Visible="false" />
                </td>
            </tr>
            <tr>
                <td class="TITLE">
                    <asp:Label runat="server" ID="lHeader" />
                    <div style="float:right">
                        <asp:Button runat="server" ID="btnNew" Text="New" CssClass="all_button" />
                    </div>
                </td>
            </tr>           
            <tr>
                <td>
                <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
                    <tr>
                        <td width="15%" class="LabelTD"><asp:Label runat="server" ID="lbl_DOC_TYPE" /> </td>
                        <td><asp:label runat="server" ID="DOC_TYPE"  /></td>
                        <td width="15%" class="LabelTD"><asp:Label runat="server" ID="lbl_DOC_NO" /> </td>
                        <td>
                            <asp:Label runat="server" ID="DOC_NO" />
                        </td>
                    </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td class="TITLE"><asp:Label runat="server" id="lblHDGV" Text="Notes Listing" /></td>
            </tr>
            <tr>
                <td>
                <asp:updatepanel runat="server" ID="MainUDP" RenderMode="Inline">
                    <ContentTemplate>
                     <asp:panel runat="server" id="GVPan" ScrollBars="Auto" style="max-height: 600px">
                     <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Note Found."
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" ShowHeaderWhenEmpty="true"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="SYS_PK" HorizontalAlign="Left">
                                <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                                <Columns>
                                    <asp:TemplateField ItemStyle-Width="15%" ItemStyle-HorizontalAlign="Center">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemTemplate>
                                            <asp:Button runat="server" ID="btnView" CssClass="all_button" Text="View" CommandName="VIEW" />
                                            <asp:Button runat="server" ID="btnReply" CssClass="all_button" Text="Reply" CommandName="REPLY" />
                                            <asp:HiddenField runat="server" ID="sys_pk" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                     <asp:TemplateField HeaderText="Posted By" ItemStyle-Width="10%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="TASK_FROM" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Subject" ItemStyle-Width="25%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="TASK_SUBJECT" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Details" ItemStyle-Width="30%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="TASK_DETAILS" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Date" ItemStyle-Width="15%">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="SYS_CD" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                      </asp:GridView>                    
                    </asp:panel>
                    </ContentTemplate>
                </asp:updatepanel>
                </td>
            </tr>
        </table>


        <asp:UpdatePanel runat="server" id="NewUDP" RenderMode="Inline">
        <ContentTemplate>
            <asp:HiddenField runat="server" ID="dummy" />
            <asp:Panel ID="pnlDtl" runat="server" CssClass="modalPopup" style="display: none" ><!---->
                <table width="700px">
                   <tr>
                        <td class="TITLE">
                            <asp:Panel runat="Server" ID="PanelDrag1" Style="cursor: move;">
                            <font size="2"><asp:Label runat="server" ID="lHeaderNew" /></font>
                            </asp:Panel>
                        </td>
                   </tr>                   
                   <tr>
                      <td>
                            <table width="100%" border="0" cellspacing="1" cellpadding="1">
                                <tr>
                                    <td class="LabelTD" width="15%">
                                        <asp:Label runat="server" ID="lbl_TASK_PRIORITY" text="Piority" />
                                    </td>
                                    <td>
                                        <asp:DropDownList runat="server" ID="TASK_PRIORITY_N">
                                            <asp:ListItem Value="" Text="SELECT" />
                                            <asp:ListItem Value="H" Text="High" />
                                            <asp:ListItem Value="M" Text="Middle" />
                                            <asp:ListItem Value="L" Text="Low" />
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="LabelTD">
                                        <asp:Label runat="server" ID="lbl_TASK_SUBJECT" text="Subject" />
                                    </td>
                                    <td>
                                        <asp:label runat="server" ID="vw_TASK_SUBJECT" />
                                        <asp:textbox runat="server" ID="TASK_SUBJECT_N" width="400px" MaxLength="80" onkeydown="return (event.keyCode!=13);"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="LabelTD">
                                        <asp:Label runat="server" ID="lbl_TASK_DETAILS" text="Details" />
                                    </td>
                                    <td>
                                        <asp:label runat="server" ID="vw_TASK_DETAILS" />
                                        <asp:textbox runat="server" ID="TASK_DETAILS" width="400px" textmode="MultiLine" Rows="5" onKeyDown="limitText(this,4000);" onKeyUp="limitText(this,4000);" onBlur="limitText(this,4000);" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="LabelTD">
                                        <asp:Label runat="server" ID="lbl_TASK_FROM" text="Posted By" />
                                    </td>
                                    <td>
                                        <asp:Label runat="server" ID="TASK_FROM" />
                                    </td>
                                </tr>
                                <tr runat="server" id="TR_SYS_CD">
                                    <td class="LabelTD">
                                        <asp:Label runat="server" ID="lbl_SYS_CD" text="Post on" />
                                    </td>
                                    <td>
                                        <asp:Label runat="server" ID="SYS_CD" />
                                    </td>
                                </tr>
                                <tr runat="server" id="Alert_TO_TR">
                                    <td class="LabelTD">
                                        <asp:Label runat="server" ID="lbl_ALERT_TO" text="Alert to" />
                                    </td>
                                    <td>
                                        <asp:CheckBoxList runat="server" ID="ALERT_TO" RepeatDirection="Vertical"></asp:CheckBoxList>
                                    </td>
                                </tr>
                            </table>
                      </td>
                   </tr>
                   <tr>
                        <td class="TITLE" align="left">
                            <asp:hiddenfield runat="server" ID="CAction" />
                            <asp:Button runat="server" ID="btnPnlSave" Text="Confirm" CssClass="all_button" />
                            <asp:Button runat="server" ID="btnPnlClose" Text= "Cancel" CssClass="all_button" />
                        </td>
                   </tr>
                </table>
            </asp:Panel> 

            <asp:ModalPopupExtender ID="pnlDtl_ModalPopupExtender" runat="server"
                DynamicServicePath="" 
                Enabled="True" 
                TargetControlID="dummy" 
                PopupControlID="pnlDtl"
                BackgroundCssClass="modalBackground"
                DropShadow="true"         
                CancelControlID="btnPnlClose"
                Y="50"
                PopupDragHandleControlID="PanelDrag1" RepositionMode="None">
            </asp:ModalPopupExtender>
        </ContentTemplate>        
        </asp:UpdatePanel>
    </div>
    </form>    
</body>
</html>
