<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SALES_ITM_BALReport.aspx.vb" Inherits="SALES_ITM_BALReport" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />  

</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="form1" runat="server">             

     <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black;text-align:center">         
                        <asp:Label ID="lbl_STORER_CODE" runat="server" >Storer</asp:Label>
                  
           <asp:DropDownList ID="STORER_CODE" runat="server" Width="150px" AutoPostBack="True">
                        </asp:DropDownList>
         <br />
         <br />

          <asp:Button ID="btnSales_Report" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Sales Admin Shortage Report" OnClick="btnSales_Report_Click"/> 
          <%--OnClientClick="return confirm('Are you sure to print this?')"--%>
       &nbsp;&nbsp;
          <asp:Button ID="btnITM_BAL_Report" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Item Balance Report" OnClick="btnITM_BAL_Report_Click" />
         &nbsp;&nbsp;
          <asp:Button ID="btnSendEmail" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Send Email" /><br /><br />
         <asp:Label ID="lblMsg" runat="server" Text=""></asp:Label>
         </div>
       
    </form>
</body>
</html>
