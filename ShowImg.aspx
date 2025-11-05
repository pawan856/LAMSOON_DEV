<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ShowImg.aspx.vb" Inherits="ShowImg" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
        <asp:Image runat="server" id="mainImage" />
    </div>
    </form>
</body>
</html>
