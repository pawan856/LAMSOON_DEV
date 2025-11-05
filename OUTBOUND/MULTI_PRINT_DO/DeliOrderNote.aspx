<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DeliOrderNote.aspx.vb" Inherits="OUTBOUND_DO_DELI_ORDER_DeliOrderNote" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
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
    <script language="javascript">
        function getLoad() {
            var edi = $("#EDI_SIR_NO").val();
            if (jQuery.trim(edi).length > 0) {
                var load_modalPopup = $find('load_ModalPopupExtender');
                load_modalPopup.show();
            }
        }      
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
          <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td class="TITLE" align="left" colspan="2">
                    <asp:Label runat="server" ID="lHeader" />
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
             <td class="LabelTD" valign="top" width="30%">
                    <asp:Label runat="server" ID="lbl_EDI_SIR_NO" />：
                </td>
                <td>
                    <asp:textbox runat="server" id="EDI_SIR_NO" width="600px" CssClass="REQUIRED" />
                </td>
             </tr>
             <tr>
             <td class="LabelTD" valign="top" width="30%">
                    <asp:Label runat="server" ID="lbl_WP_NO" />：
                </td>
                <td>
                    <asp:textbox runat="server" id="WP_NO" width="600px" CssClass="REQUIRED" />
                </td>
             </tr>
             <tr>
                <td class="TITLE" align="right" colspan="2">
                    <asp:button runat="server" ID="btnGen" CssClass="all_button" OnClientClick="javascript:getLoad();" />
                </td>
             </tr>
          </table>        
           <iframe id="iframe" src="" runat="server" width="0px" height="0px"></iframe>
        <asp:HiddenField ID="loadDummy" runat="server" />
             <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
                ScrollBars="None">
                <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
                    height: 80px">
                    <tr>
                        <td align="center" style="background-color: White; width: 100%; height: 80px; vertical-align: middle">                            
                            <font color="#193B65" style="width: 100%; text-align: center; font-size: 16px;">Generating...please wait</font>
                        </td>
                    </tr>                    
                </table>
            </asp:Panel>
            <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
                TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
                DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
                Y="250">
            </asp:ModalPopupExtender>
    </div>

  
    </form>
</body>
</html>