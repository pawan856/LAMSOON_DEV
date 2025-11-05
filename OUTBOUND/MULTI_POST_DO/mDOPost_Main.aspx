<%@ Page Language="VB" AutoEventWireup="false" CodeFile="mDOPost_Main.aspx.vb" Inherits="OUTBOUND_MULTI_DO_mDOPost_Main" MaintainScrollPositionOnPostback="true" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Multiple DO</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script src="../../js/jquery-1.9.1.min.js"></script>
<script src="../../js/jquery-migrate-1.1.1.min.js"></script> 
<div runat="server" id="DIVSCRIPT">
<script language="javascript">
    function maskKey(objEvent) {
        var iKeyCode;
        iKeyCode = objEvent.keyCode;
        if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
        return false;
    }

    function maskDate(objEvent) {
        var iKeyCode;
        iKeyCode = objEvent.keyCode;
        if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
        return false;
    }

    function SelectAll() {
        var frm = document.forms[0];
        for (i = 0; i < frm.elements.length; i++) {
            if (frm.elements[i].type == "checkbox") {
                frm.elements[i].checked = true;
            }
        }
    }
    function UnSelectAll() {
        var frm = document.forms[0];
        for (i = 0; i < frm.elements.length; i++) {
            if (frm.elements[i].type == "checkbox") {
                frm.elements[i].checked = false;
            }
        }
    }   
    
</script>
</div>
 <style type="text/css">
        .overlay
        {
          position: fixed;
          z-index: 98;
          top: 0px;
          left: 0px;
          right: 0px;
          bottom: 0px;
          background-color: #aaa;          
          opacity: 0.9;
        }
        .overlayContent
        {
          z-index: 99;
          margin: 250px auto;
          width: 250px;
          height: 250px;          
          text-align: center;
        }
        .overlayContent h2
        {
            font-size: 18px;
            font-weight: bold;
            color: #000;
        }
        .overlayContent img
        {
          width: 80px;
          height: 80px;          
        }
    </style>

</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release"  AsyncPostBackTimeout="3600" />
    <input type="hidden" name="moduleAction" value=""/>
    <table border="0" cellspacing="1" cellpadding="1" align="center" width="640px">
            <tr>
                <td class="TITLE" colspan="2">
                    <asp:Label runat="server" ID="lheader" />                
                </td>
            </tr>
            <tr>
                <td class="LabelTD" valign="top" width="30%">
                    <asp:Label runat="server" ID="lbl_STORER_CODE" />：
                </td>
                <td>
                    <asp:DropDownList runat="server" id="STORER_CODE" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" valign="top" width="15%">
                    <asp:Label runat="server" ID="lbl_DO_EDI_SIR_NO" />：
                </td>
                <td>
                    <asp:textbox runat="server" ID="DO_EDI_SIR_NO" />
                </td>
            </tr>

        <tr>
                <td class="LabelTD" valign="top" width="15%">
                    <asp:Label runat="server" ID="lbl_SO_NO" />：
                </td>
                <td>
                    <asp:textbox runat="server" ID="SO_NUMBER" />
                </td>
            </tr>

            <tr>
                <td class="LabelTD" valign="top" width="15%">
                    <asp:Label runat="server" ID="lbl_DO_DATE" />：
                </td>
                <td>
                    From : <asp:TextBox ID="DO_DATE_FR" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_DATE_FR" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                    To: <asp:TextBox ID="DO_DATE_TO" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_DATE_TO" PopupButtonID="btnDATE_ID2" Format="dd/MM/yyyy" />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:button runat="server" ID="btnSearch" CssClass="all_button" /><!>
                    <asp:button runat="server" ID="btnReset" CssClass="all_button" OnClientClick="Javascript:window.location='./mDOPost_Main.aspx'; return false;" />
                </td>
            </tr>
    </table>
    <br />
    <asp:UpdatePanel runat="server" ID="udp1" RenderMode="Inline">
    <ContentTemplate>
            <asp:Panel runat="server" ID="GVPNL" Visible="false">
                <table border="0" cellspacing="1" cellpadding="1" align="center" width="1024px">  
                     <tr>
                        <td class="TITLE" >
                            <asp:Label runat="server" ID="lblGV" />
                        </td>
                     </tr>
                     <tr>
                        <td class="TITLE" valign="top" align="center" >
                            <div align="left">Post Result:</div><asp:TextBox runat="server" ID="errText" TextMode="MultiLine" Rows="10" Width="95%" />                       
                        </td>
                     </tr>
                     <tr>
                        <td class="menuTD">    
                            <asp:Button runat="server" ID="btnSelAll" CssClass="all_button" OnClientClick="Javascript:SelectAll();return false" /><!>
                            <asp:Button runat="server" ID="btnUnSel" CssClass="all_button" OnClientClick="Javascript:UnSelectAll();return false" /><!>

                            <div style="float:right;">
                                <asp:Button runat="server" ID="BtnGen" CssClass="all_button" />
                                <asp:ConfirmButtonExtender runat="server" ID="cfm1" TargetControlID="btnGen" ConfirmText="New DOs will be generated from selected CO, Confirm to proceed?" />
                            </div>
                        </td>
                     </tr>
                     <tr>
                        <td>
                            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                    Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="3" CaptionAlign="Top" DataKeyNames="DO_CODE" HorizontalAlign="Left">
                                       <RowStyle CssClass="GV" />
                                    <Columns>
                                         <asp:TemplateField ItemStyle-HorizontalAlign="Center">
                                            <ControlStyle Width="30px" />
                                            <HeaderStyle HorizontalAlign="Left" Width="20px" />
                                            <ItemTemplate>
                                                <asp:CheckBox runat="server" ID="checkYN" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Organizations">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label runat="server" ID="sto_name" Font-Size="11px" />
                                                <asp:hiddenfield ID="STORER_CODE" runat="server"/>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="DO Code">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="DO_CODE" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Customer Order No.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="DO_EDI_SIR_NO" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>                         
                                        <asp:TemplateField HeaderText="DO Date">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="DO_DATE" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                         <%--add new field route--%>
                                         <asp:TemplateField HeaderText="Route">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="ROUTE_ID" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <%--end--%>
                                        <asp:TemplateField HeaderText="SO No.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="SO_NUMBER" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Customer Name">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="CUS_NAME" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Status">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemTemplate>
                                                <asp:Label ID="DO_STATUS" runat="server" Font-Size="11px"></asp:Label>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                    </Columns>                        
                                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                             </asp:GridView>    
                        </td>
                     </tr>
                     <tr>
                        <td class="menuTD">    
                            <div style="float:right;">
                                <asp:Button runat="server" ID="btnGen2" CssClass="all_button" />
                                <asp:ConfirmButtonExtender runat="server" ID="cfm2" TargetControlID="btnGen2" ConfirmText="New DOs will be generated from selected CO, Confirm to proceed?" />
                            </div>
                        </td>
                     </tr>
                </table>
                </asp:Panel>
    </ContentTemplate>
    </asp:UpdatePanel>
    <asp:UpdateProgress runat="server" ID="UpdateProgress1" DisplayAfter="0" AssociatedUpdatePanelID="udp1">
    <ProgressTemplate> 
            <div class="overlay" />
            <div class="overlayContent">
                <h2>Loading...Please Wait...</h2>
                <img src="../../images/ajax-loader.gif" alt="Loading" border="1" />
            </div>
    </ProgressTemplate>
    </asp:UpdateProgress>
    </div>    
    </form>
</body>
</html>
