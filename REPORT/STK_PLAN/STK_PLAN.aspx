<%@ Page Language="VB" AutoEventWireup="false" CodeFile="STK_PLAN.aspx.vb" Inherits="OPERATION_SCHK_CheckList" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>        
        <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
        </asp:ToolkitScriptManager>
        <input type="button" value="Back" class="all_button" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=RPT_STK_PLAN'" />    
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" 
            Font-Size="8pt" height="800px" style="position: static" width="1200px"
            WaitMessageFont-Names="Verdana" WaitMessageFont-Size="14pt">
            <LocalReport ReportPath="REPORT\STK_PLAN\STK_PLAN.RDLC">
            </LocalReport>
        </rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>
