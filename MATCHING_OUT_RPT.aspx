<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MATCHING_OUT_RPT.aspx.vb" Inherits="MATCHING_OUT_RPT" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>


<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">    
    <br />
    <form id="form1" runat="server">  
        <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
        <input type="hidden" name="moduleAction" value="" />

           <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black;text-align:center"> 
          <asp:Button ID="btnDownloadExcel" runat="server" BackColor="black" Font-Bold="True" ForeColor="White" Text="Download" OnClick="btnDownloadExcel_Click"/> 
         </div>

         <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>
       
    </form>
</body>

</html>
