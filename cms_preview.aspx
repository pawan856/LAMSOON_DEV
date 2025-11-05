<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cms_preview.aspx.vb" Inherits="cms_preview"  %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <script src="js/jquery-1.9.1.min.js"></script>
    <script src="js/jquery-migrate-1.1.1.min.js"></script> 

</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />
        <div>    
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="970px" Height="700px" KeepSessionAlive="true" ZoomMode="PageWidth" ZoomPercent="100" cssclass="ReportViewerControl">
        </rsweb:ReportViewer>
        </div>
    </form>    
</body>
</html>

<script type="text/javascript" language="javascript">
    $(window).unload(function () {
        PageMethods.deleteTempFiles(function (returnValue) {                        
        });
    }); 

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