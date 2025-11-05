<%@ Page Language="VB" AutoEventWireup="false" CodeFile="stock_receipt.aspx.vb" Inherits="INBOUND_GR_GR_STOCK_REP_stock_receipt" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link rel="stylesheet" href="../../../stylesheet/button.css" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <table cellpadding="0" cellspacing="0" border="0">
            <tr>
                <td>
                    <asp:Button runat="server" ID="btnSort" Text="Sort By Carton No." CssClass="all_button" Visible="false" />
                </td>
            </tr>
            <tr>
                <td>
                    <rsweb:ReportViewer ID="ReportViewer1" runat="server" font-size="8pt"
            height="750px" style="position: static" width="1100px"></rsweb:ReportViewer>       
                </td>
            </tr>
        </table>
         
    </div>
    </form>
</body>
</html>

<script type="text/javascript" language="javascript">    
        //var obj=document.getElementById('ReportViewer1_ctl01_ctl05_ctl00');    
        //if (obj)    
        //{ var element_length=obj.length;     
        //var index=-1;        
        //for(z=0;z<element_length;z++)      
        //{           
        // if (obj.options[z].value=='Excel')    
        // {index=z;}                    
        //}    
        //if(index!=-1)       
        //obj.remove(index);
        //}       
 </script>   

