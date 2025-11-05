	<%@ Page Language="VB" AutoEventWireup="false" CodeFile="main.aspx.vb" Inherits="main" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="content-type" content="text/html; charset=<%=G_HTML_LANG_ENCODING%>">
    <title>WMS v3.02</title>
</head>
    
<frameset rows="140,*" cols="*" border="0" framespacing="0">
	<!--<frame src="logon.top.asp?ranDate=<%=Server.URLEncode(now)%>" name="top" scrolling="NO" marginwidth="0" marginheight="0" noresize frameborder="NO">-->
	<frame src="menu.aspx?ranDate=<%=Server.URLEncode(now)%>" name="top" scrolling="NO" marginwidth="0" marginheight="0" noresize frameborder="NO">
	<frame name="main" id="main" marginwidth="0" marginheight="0" noresize frameborder="NO" runat="server" />
	
<!--	<frame src="printFrame.html" name="printhere" marginwidth="0" marginheight="0" noresize frameborder="NO">-->
</frameset>
<noframes>
<body bgcolor="#FFFFFF">
</body>
</noframes>
</html>
