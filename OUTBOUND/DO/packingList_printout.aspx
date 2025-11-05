<%@ Page Language="VB" AutoEventWireup="false" CodeFile="packingList_printout.aspx.vb" Inherits="OUTBOUND_DO_packingList_printout" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
   <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="600px" 
            Width="1150px">
        </rsweb:ReportViewer>
    
    </div>
    </form>
</body>
</html>
