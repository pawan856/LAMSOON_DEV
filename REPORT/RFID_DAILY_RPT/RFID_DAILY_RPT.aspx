<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RFID_DAILY_RPT.aspx.vb" Inherits="RFID_DAILY_RPT" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
    <input id="btnBack2" runat="server" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=RFID_DAILY_RPT'"
         class="all_button" /><br />
         <rsweb:ReportViewer ID="ReportViewer1" runat="server" font-size="8pt"
            height="768px" style="position: static" width="800px"></rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>