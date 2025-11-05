<%@ Page Language="VB" AutoEventWireup="false" CodeFile="main_news.aspx.vb" Inherits="main_news" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="ajx" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="content-type" content="text/html; charset=<%=G_HTML_LANG_ENCODING%>">
    <title>LOGON</title>
    <link rel="stylesheet" href="stylesheet/general.css" type="text/css" />
    <link rel="stylesheet" href="stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="stylesheet/main_page_text.css" type="text/css">
    <style type="text/css">

 p.MsoNormal
	{margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	        margin-left: 0in;
            margin-right: 0in;
            margin-top: 0in;
        }
    </style>
    <script language="javascript">
      
    </script>
</head>
<body bgcolor="#FFFFFF" link="#003366" vlink="#666666" alink="#FF6600">
<br />
<br />
<br />
    <form id="form1" runat="server">
    <div>
    <ajx:ToolkitScriptManager runat="server" ID="scriptman1"  />
    <table >
    <tr>
        <td class="main_page_text" style="background-color:White;">
            <p class="MsoNormal" style="margin-bottom:6.0pt">
                <b><span style="font-family:
Arial;color:#193972">Welcome to Warehouose Management System (WMS)<o:p></o:p></span></b></p>
            <p class="MsoNormal" style="margin-bottom:6.0pt">
                <b><span style="font-size:9.0pt;
font-family:Arial;color:#193972"><o:p>&nbsp;</o:p></span></b></p>
            <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                <span 
                    style="font-size:9.0pt;font-family:Arial;color:#193972">
                <o:p>
                <b><span style="font-size:9.0pt;font-family:Arial;
color:#193972">eS3</span></b> is a system by Elogistics Hong Kong Limited for the management of 
                small to medium size logistics operations including Warehouse Management, 
                Transportation Management and Order Management. The system is completely 
                web-based with option for client-server version depending on customer 
                requirements. It belongs to a suite of logistics and product sourcing software 
                solutions covering the whole product life-cycle from order placement through to 
                a manufacturing and order tracking and to shipment and warehouse management.</o:p></span></p>
            <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                <b><span style="font-size:9.0pt;font-family:Arial;
color:#193972"><span style="font-size:9.0pt;
font-family:Arial;color:#193972;mso-bidi-font-weight:bold"> <o:p>Elogistics Hong Kong Limited</o:p></span></span></b><span style="font-size:9.0pt;
font-family:Arial;color:#193972;mso-bidi-font-weight:bold"><o:p> is a logistics solution 
                provider focused on providing systems for Hong Kong and other parts of China.</o:p></span></p>
            
            <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                <b>
                <span style="font-size:9.0pt;font-family:Arial;mso-fareast-font-family:PMingLiU;
color:#193972;mso-ansi-language:EN-US;mso-fareast-language:ZH-TW;mso-bidi-language:
AR-SA">
                Warehouse Management </span></b>     
            </p>
                        <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                            <b>
                            <span style="font-size:9.0pt;font-family:Arial;mso-fareast-font-family:PMingLiU;
color:#193972;mso-ansi-language:EN-US;mso-fareast-language:ZH-TW;mso-bidi-language:
AR-SA">
                            Transportation Management </span></b>     
            </p>
                        <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                            <b>
                            <span style="font-size:9.0pt;font-family:Arial;mso-fareast-font-family:PMingLiU;
color:#193972;mso-ansi-language:EN-US;mso-fareast-language:ZH-TW;mso-bidi-language:
AR-SA">
                            Product Sourcing and Quotation Management </span></b>     
            </p>
                        <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                            <b>
                            <span style="font-size:9.0pt;font-family:Arial;mso-fareast-font-family:PMingLiU;
color:#193972;mso-ansi-language:EN-US;mso-fareast-language:ZH-TW;mso-bidi-language:
AR-SA">
                            Merchandising and Quality Management </span></b>     
            </p>
                        <p class="MsoNormal" style="margin-top:0in;margin-right:0in;margin-bottom:6.0pt;
margin-left:.25in">
                            <b>
                            <span style="font-size:9.0pt;font-family:Arial;mso-fareast-font-family:PMingLiU;
color:#193972;mso-ansi-language:EN-US;mso-fareast-language:ZH-TW;mso-bidi-language:
AR-SA">
                            Order Tracking and Order Alerts Management</span></b></p>
        </td>
        <td valign="top" style="background-color:White;">
            <img src="images/main_news.JPG" />
        </td>
    </tr>
    <tr>
        <td colspan="2" style="background-color:White;">
            <asp:Button runat="server" ID="btnToDo" Text="Show TO-DO List" CssClass="all_button" />
        </td>
    </tr>
    </table>

    <asp:panel runat="server" id="contentPnl" style="display:none;">
        <table width="640px">
            <tr>
                <td style="background-color:#ffcc66;" colspan="3">
                    <asp:Panel runat="server" ID="dragPanel" Width="100%" style="cursor:move;">
                    <asp:Label runat="server" ID="Dragtitle" Text="To-Do List" ForeColor="White"  Font-Size="Medium" Font-Bold="true" />
                    </asp:Panel>                    
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:LinkButton runat="server" id="alertLink" Font-Size="16px" ForeColor="#D14383" Text="You have 0 Alert." OnClientClick="javascript:window.open('./ALERT/AlertMain.aspx','_blank', 'width=1200,height=600,location=0, left=10, top=10');return false;"  />
                </td>
            </tr>
            <asp:Panel runat="server" ID="pnlHeader">
            <tr>
                <td style="background-color:#ffcc66;">
                    Unit
                </td>
                <td style="background-color:#ffcc66;">
                    Item
                </td>
                <td style="background-color:#ffcc66;">
                    Count
                </td>
            </tr>  
            </asp:Panel>
            <asp:panel  runat="server" id="RCV_PNL" Visible="false">
            <tr>
                <td rowspan="2" valign="bottom">Receiving Unit</td>
                <td><asp:LinkButton runat="server" ID="BTN_PO_ALC_COUNT" text="Outstanding PO" /></td>
                <td align="right"><asp:Label runat="server" id="PO_ALC_COUNT" /></td>
            </tr>          
            <tr>
                <td><asp:LinkButton runat="server" ID="BTN_SR_ALC_COUNT" text="Pending Credit Return" /></td>
                <td align="right"><asp:Label runat="server" id="SR_ALC_COUNT" /></td>
            </tr>
            </asp:panel>
            <asp:panel  runat="server" id="ISS_PNL" Visible="false">
            <tr >
                <td rowspan="2" valign="bottom">Issuing Unit</td>
                <td><asp:LinkButton runat="server" ID="BTN_SIR_ALC_COUNT" text="Outstanding CO" /></td>
                <td align="right"><asp:label runat="server" id="SIR_ALC_COUNT" text="1" /></td>
            </tr>          
            <tr>
                <td><asp:LinkButton runat="server" ID="BTN_WIT_ALC_COUNT" text="Pending DO" /></td>
                <td align="right"><asp:label runat="server" id="WIT_ALC_COUNT" /></td>
            </tr>
            </asp:panel>
            <asp:panel  runat="server" id="LMA_PNL" Visible="false">
            <tr>
                <td rowspan="4" valign="bottom">Lamma Unit</td>
                <td><asp:LinkButton runat="server" ID="BTN_PO_LMA_COUNT" text="Outstanding PO (LMA)" /></td>
                <td align="right"><asp:label runat="server" id="PO_LMA_COUNT" /></td>
            </tr>      
            <tr>
                <td><asp:LinkButton runat="server" ID="BTN_SR_LMA_COUNT" text="Pending Credit Return (LMA)" /></td>
                <td align="right"><asp:label runat="server" id="SR_LMA_COUNT" /></td>
            </tr>   
            <tr>
               <td><asp:LinkButton runat="server" ID="BTN_SIR_LMA_COUNT" text="Outstanding CO (LMA)" /></td>
                <td align="right"><asp:label runat="server" id="SIR_LMA_COUNT" /></td>
            </tr>          
            <tr>
                <td><asp:LinkButton runat="server" ID="BTN_WIT_LMA_COUNT" text="Pending DO (LMA)" /></td>
                <td align="right"><asp:label runat="server" id="WIT_LMA_COUNT" /></td>
            </tr> 
            </asp:panel>            
            <tr runat="server" id="emptyTR" visible="false">
                <td colspan="3">&nbsp;</td>                
            </tr>
            <tr>
                <td class="TITLE" colspan="3">
                    <asp:button runat="server" ID="btnClose" Text="Close" CssClass="all_button" />
                </td>
            </tr>
          </table>
    </asp:panel>
    <asp:HiddenField runat="server" ID="dummy" />
    <ajx:ModalPopupExtender ID="ToDo_POPUP" runat="server" 
    DynamicServicePath="" 
    Enabled="True" 
    TargetControlID="dummy" 
    PopupControlID="contentPnl"
    BackgroundCssClass="modalBackground"
    DropShadow="true"         
    CancelControlID="btnClose"
    Y="50"    
    PopupDragHandleControlID="dragPanel" >
</ajx:ModalPopupExtender>
    </div>
    </form>
</body>
</html>
