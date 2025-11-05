<%@ Page Language="VB" AutoEventWireup="false" CodeFile="eoffice_login.aspx.vb" Inherits="eoffice_login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Warehouse Management System WMS Version 3.02</title>
    <link rel="stylesheet" href="stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="stylesheet/button.css" type="text/css">
    <style type="text/css">
        .textfield {
            font-family: "Arial", "Helvetica", "sans-serif";
            font-size: 12px;
        }

        .putdnmenu {
            font-family: "Arial", "Helvetica", "sans-serif";
            font-size: 12px;
            background-color: #dddddd;
        }

        .NEWTEXT {
            font-family: "Arial", "Helvetica", "sans-serif";
            font-size: 12px;
            text-decoration: none;
        }

        .style2 {
            width: 416px;
        }
    </style>

    <script type="text/javascript" language="javascript">
        function ForwardFrameToLogin() {
            if (top.location.href != window.location.href) {
                top.location.href = window.location.href;
            }
        }
    </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="ForwardFrameToLogin();document.forms[0].Login1_UserName.focus();">
    <p>
        &nbsp;
    </p>
    <form method="post" name="myform" id="myform" onsubmit="upperCaseResults();//upperCasePassword();"
        runat="server">
        <div>
            <br />
        </div>
        <table border="0" cellspacing="0" cellpadding="0" align="center" width="757">
            <tr>
                <td>
                    <img src="images/version/login_1.gif" />
                </td>
                <td>&nbsp;
                </td>
                <td>
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td colspan="2">
                                <img src="mis_login_images/login_2.bmp" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <img src="mis_login_images/login_logo.png" />
                            </td>
                            <td background="mis_login_images/login_5.png" valign="top">
                                <table cellspacing="0" cellpadding="1" border="0">
                                    <tbody>
                                        <tr>
                                            <td></td>
                                        </tr>
                                        <tr>
                                            <td align="left" class="newtext">
                                                <font size="2">
                                                    <asp:Label ID="Label1" runat="server" AssociatedControlID="SelConn">DataBase</asp:Label>
                                                </font><span>&nbsp;:</span>
                                                <asp:DropDownList Width="168px" runat="server" ID="SelConn">
                                                    <asp:ListItem Value="Production" Text="Production" />
                                                    <asp:ListItem Value="Archived" Text="Archived" />
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="style2">
                                                <asp:Login ID="Login1" runat="server" BorderPadding="0" BorderStyle="Solid" BorderWidth="0px"
                                                    DisplayRememberMe="False" LoginButtonText="Login" TitleText="" FailureText="Login Failed. Please try again.">
                                                    <TextBoxStyle Font-Size="0.8em" />
                                                    <LoginButtonStyle Font-Names="Arial" CssClass="all_button" />
                                                    <LayoutTemplate>
                                                        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; padding-left: 30px">
                                                            <tr>
                                                                <td>
                                                                    <table border="0" cellpadding="1">

                                                                        <tr>
                                                                            <td align="left" class="newtext" width="70px">
                                                                                <font size="2">
                                                                                    <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="UserName">User ID</asp:Label>
                                                                                </font>
                                                                            </td>
                                                                            <td width="5px" class="newtext">
                                                                                <font size="2">:</font>
                                                                            </td>
                                                                            <td width="180px" nowrap>
                                                                                <asp:TextBox ID="UserName" runat="server" onBlur="upperCaseResults()" Style="font-family: 'Arial'"
                                                                                    Text="" Width="160px"></asp:TextBox>
                                                                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="UserName"
                                                                                    ErrorMessage="User Name is required." ToolTip="User ID is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td align="left" class="newtext" width="70px">
                                                                                <font size="2">
                                                                                    <asp:Label ID="PasswordLabel" runat="server" AssociatedControlID="Password">Password</asp:Label>
                                                                                </font>
                                                                            </td>
                                                                            <td width="5px" class="newtext">
                                                                                <font size="2">:</font>
                                                                            </td>
                                                                            <td width="180px" nowrap>
                                                                                <asp:TextBox ID="Password" runat="server" tyle="font-family:'Arial'" Text="" Width="160px"
                                                                                    TextMode="Password"></asp:TextBox>
                                                                                <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" ControlToValidate="Password"
                                                                                    ErrorMessage="Password is required." ToolTip="Password is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                                                            </td>
                                                                        </tr>

                                                                        <tr>
                                                                            <td colspan="2">&nbsp;
                                                                            </td>
                                                                            <td width="100%" align="left" style="padding: 1px">
                                                                                <asp:Button ID="LoginButton" runat="server" CommandName="Login" CssClass="all_button"
                                                                                    Font-Names="Arial" Text="Login" Width="80px" ValidationGroup="Login1" />
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td colspan="2">&nbsp;
                                                                            </td>
                                                                            <td style="color: Red;" nowrap>
                                                                                <font size="3">
                                                                                    <asp:Literal ID="FailureText" runat="server" EnableViewState="False"></asp:Literal>
                                                                                </font>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </LayoutTemplate>
                                                    <InstructionTextStyle Font-Italic="True" ForeColor="Black" />
                                                    <TitleTextStyle BackColor="#5D7B9D" Font-Bold="True" Font-Size="0.9em" ForeColor="White" />
                                                </asp:Login>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>

                                <asp:DropDownList runat="server" ID="SelLang">
                                    <asp:ListItem Value="E" Text="English" />
                                    <asp:ListItem Value="C" Text="繁體中文" />
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
<script language="javascript">
    function upperCaseResults() {
        var newResults = document.myform.Login1_UserName.value
        document.myform.Login1_UserName.value = newResults.toUpperCase()

    }

    function upperCasePassword() {
        var newResults = document.myform.password.value
        document.myform.password.value = newResults.toUpperCase()

    }


    function StripLeadSpace(InString) {
        var i = 0;

        for (i = 0; InString.substring(i, i + 1) == " "; i++);

        return (InString.substring(i, InString.length));
    }

    function CheckForm() {
        var nErr = 0;
        var vUserid = "";
        var vPassword = "";

        vUserid = StripLeadSpace(myform.Login1_UserName.value);
        vPassword = StripLeadSpace(myform.password.value);

        if (vUserid == "" && vPassword == "") {
            alert("User ID and Password cannot be empty");
            return false;
        }

        if (vUserid == "") {
            alert("User ID cannot be empty");
            return false;
        }
        if (vPassword == "") {
            alert("Password cannot be empty");
            return false;
        }
        return true;
    }

    window.focus()
</script>
