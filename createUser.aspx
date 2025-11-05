<%@ Page Language="VB" AutoEventWireup="false" CodeFile="createUser.aspx.vb" Inherits="createUser" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        ID:<asp:TextBox runat="server" ID="user_id" />
        PWD:<asp:TextBox runat="server" id="pwd" />
        <asp:Button runat="server" Text="Add Admin" ID="btnADD"/>
    </div>
    </form>
</body>
</html>
