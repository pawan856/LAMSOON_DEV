<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sitemap.aspx.vb" Inherits="sitemap" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="content-type" content="text/html; charset=<%=G_HTML_LANG_ENCODING%>">
<title>Integrated Merchandising System</title>

<link rel="stylesheet" href="stylesheet/newtext.css" type="text/css">
<link rel="stylesheet" href="stylesheet/button.css" type="text/css">
<link rel="stylesheet" href="stylesheet/main_page_text.css" type="text/css">
<link rel="stylesheet" href="stylesheet/main_menu.css" type="text/css">
<script language="JavaScript">

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

</script>

<style type="text/css">

.menu_text {  font-family: "Arial", "Helvetica", "sans-serif"; font-size: 11px; font-weight: bold}

</style>
</head>
<body bgcolor="#FFFFFF" link="#003366" vlink="#666666" alink="#FF6600" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="menuform" runat="server">
    </form>
</body>
</html>
