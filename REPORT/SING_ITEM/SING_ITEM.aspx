<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sing_item.aspx.vb" Inherits="RPT_SING_ITEM" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>
    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
         <rsweb:ReportViewer ID="ReportViewer1" runat="server" font-size="8pt"
            height="768px" style="position: static" width="1100px"></rsweb:ReportViewer>
    </div>
    </form>
</body>
</html>

<%--<script type="text/javascript" language="javascript">    
    var obj=document.getElementById('ReportViewer1_ctl01_ctl05_ctl00');    
    if (obj)    
    { var element_length=obj.length;     
    var index=-1;        
    for(z=0;z<element_length;z++)      
    {           
     if (obj.options[z].value=='Excel')    
     {index=z;}                    
    }    
    if(index!=-1)       
    obj.remove(index);
    }       
 </script>--%>