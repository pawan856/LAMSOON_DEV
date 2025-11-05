<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cronjobImplement.aspx.vb" Inherits="cronjobImplement" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

     <script type="text/javascript">
        function ShowProgress() { 
            //var dialog = confirm("Are you sure to Import this data?");            
            //if (dialog == true) {
                setTimeout(function () {
                    var modal = $('<div />');
                    modal.addClass("modal");
                    $('body').append(modal);
                    var loading = $(".loading");
                    loading.show();
                }, 200);
            //}        
           
            //return dialog;            
            
        }
        
        
       </script>
</head>
<body onload="ShowProgress()">
    <form id="form1" runat="server">
          <%@ Reference VirtualPath="~/IMPORTEXPORT.aspx" %>
        
         <div class="loading" align="center" style="display:none">
            Loading. Please wait.<br />
            <br />
            <img src="images/ajax-loader.gif" alt="Loading..." />
        </div>          

    <div>
       <p style="margin-left: 2px;">
                    <asp:Literal ID="lblMSG" runat="server"></asp:Literal>
                </p>
    </div>
    </form>
</body>
</html>
