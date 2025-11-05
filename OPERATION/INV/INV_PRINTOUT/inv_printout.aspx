<%@ Page Language="VB" AutoEventWireup="false" CodeFile="inv_printout.aspx.vb" Inherits="OPERATION_INV_INV_PRINTOUT_inv_printout" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
    
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" font-size="8pt"
            height="700px" style="position: static" width="100%"></rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
