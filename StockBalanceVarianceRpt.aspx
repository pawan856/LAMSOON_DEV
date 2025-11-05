<%@ Page Language="VB" AutoEventWireup="false" CodeFile="StockBalanceVarianceRpt.aspx.vb" Inherits="StockBalanceVarianceRpt" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
   <title></title>
      <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
    <style>
        .DtlLabel {
    font-family: "Arial" , "Helvetica" , "sans-serif" !important;
    font-size: 12px !important;
    color: #333333 !important;
    background-color: #B6BCC4 !important ;
    font-weight: normal !important;
}
    </style>
  
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">    
    <br />
    <form id="form1" runat="server">  
        <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
        <input type="hidden" name="moduleAction" value="" />

           <div style="height: auto; width: auto; font-weight: bold; border: 0px solid black;text-align:center"> 
          <asp:Button ID="btnExportExcel" runat="server" BackColor="black" Visible="false" Font-Bold="True" ForeColor="White" Text="Export To Excel" OnClick="btnExportExcel_Click"/> 
          <br />
           <br />

               <asp:GridView ID="GridTableData" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="True" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" ShowFooter="true">
                        <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
         </div>

        <asp:HiddenField ID="RowsCount" runat="server" />    

         <!-- Update panel that decided for validation alert -->
        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
            <ContentTemplate />
        </asp:UpdatePanel>
       
    </form>
</body>

</html>
