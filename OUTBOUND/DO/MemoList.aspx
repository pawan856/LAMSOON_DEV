<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MemoList.aspx.vb" Inherits="OUTBOUND_DO_MemoList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
    <script language="javascript" src="../../js/validation.js"></script>
    <script language="javascript" src="../../js/JS_Calendar.js"></script>
    <script language="javascript" src="../../js/listUtil.js"></script>
    <script language="javascript" src="../../js/formatUtil.js"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js"></script>
    <script language="javascript">
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
    <tr>
        <td class="LabelTD" nowrap>
            <asp:Label ID="lbl_db" runat="server" />
        </td>
        <td>
            <asp:Label ID="DB_USR" runat="server" />
        </td>
        <td class="LabelTD" nowrap>
            <asp:Label ID="lbl_usr" runat="server" />
        </td>
        <td colspan="1">
            <asp:Label ID="usr_id" runat="server" />
        </td>
    </tr>
    <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
        <td colspan="4">
            <asp:TextBox ID="tb_cmd" runat="server" Height="300" Width="600" TextMode="MultiLine" />
        </td>
    </tr>
    <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
        <td class="menuTD" align="left">
            <asp:Button ID="btnRoll" CssClass="all_button" runat="server" Text="Rollback Run" />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Button ID="btnCommit" CssClass="all_button" runat="server" Text="Commit Run" />
        </td>
    </tr>
    </table>
    <br />
    <br />
    
    <div id="dtdiv" runat="server"></div>
    </div>
    </form>
</body>
</html>
