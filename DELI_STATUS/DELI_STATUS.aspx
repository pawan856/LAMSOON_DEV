<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DELI_STATUS.aspx.vb" Inherits="DELI_STATUS_DELI_STATUS" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Delivaery Status</title>
 <style type="text/css">
  table.mainTable 
  {
    border-style: solid;
    border-width: 1px;    
    
  }   
  
  .mainTable td
  {
   padding:3px;   
   font-size:12px;
  }
  
  
  td.dateBar
  {   
    font-size:18px;
    color: Black;
    background-color: #8DD8F2;   
    padding:3px;    
  }   
     
 </style>
<script language="javascript" type="text/javascript" src="../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">
    
</script>
</head>
<body style="background-color:#E6E9F0;">
    <form id="form1" runat="server">
    <div>
    <h2>
        <asp:label runat="server" ID="lblTitle" runat="server" Text="Delivery Status" />:
    </h1>
    
    <asp:Table runat="server" ID="statsTable" HorizontalAlign="Center" Width="800px" CssClass="mainTable" >
        
    </asp:Table>

    <asp:Label runat="server" ID="lblWarn" visible="false"/>
    </div>
    </form>
</body>
</html>
