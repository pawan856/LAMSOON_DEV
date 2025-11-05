<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RPT_SL_RCD.aspx.vb" Inherits="RPT_SLCABLE_RCD_main" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Download Report</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
    
    <script type="text/javascript" language="javascript" src="<%=ROOT_PATH%>js/validation.js"></script>
    <script type="text/javascript" language="javascript" src="<%=ROOT_PATH%>js/JS_Calendar.js"></script>
    <script type="text/javascript" language="javascript" src="<%=ROOT_PATH%>js/listUtil.js"></script>
    <script type="text/javascript" language="javascript" src="<%=ROOT_PATH%>js/formatUtil.js"></script>
    <script type="text/javascript" language="javascript" src="<%=ROOT_PATH%>js/formPostInterfacing.js"></script>
    <script type="text/javascript" language="javascript">
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 600px">
    <tr>
        <td class="TITLE" colspan="4">
            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td class="TITLE" colspan ="2">
                        <b><asp:Label ID="lheader" runat="server" text="Short Length Cable Detail Report" /></b>
                    </td>
                 </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td colspan="4" class="menuTD">
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td class="menuTD" width="50%">
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='<%=ROOT_PATH%>cms_search.aspx?menu_code=<%=FUN_CODE%>'"
                        class="all_button" />
                </td>
                <td class="menuTD" style="text-align:right">&nbsp;</td>
            </tr>
            </table>
        </td>
    </tr>
	<tr>
        <td class="LabelTD" style="width: 180px">
            <asp:Label ID="lbl_status" runat="server" Text="Status:" />
        </td>
        <td colspan="3">
            <asp:Label ID="dsp_status" runat="server" text="Now Loading ..."/>
        </td>
    </tr>
    <tr runat="server" id="tr_download">
		<td class="LabelTD">File Preview/Download :</td>
		<td colspan="3">
            <asp:ImageButton ID="ibtn_download" runat="server" ImageUrl="~/images/report_download.gif" Visible="False" OnClientClick="if (document.forms[0].hdf_file_path.value=='') return false; else return true;" />
        </td>
	</tr>
    <tr>
        <td colspan="4" class="menuTD">
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td class="menuTD" width="50%">
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='<%=ROOT_PATH%>cms_search.aspx?menu_code=<%=FUN_CODE%>'"
                        class="all_button" />
                </td>
                <td class="menuTD" style="text-align:right">&nbsp;</td>
            </tr>
            </table>
        </td>
    </tr>
    </table>
    </div>
    <asp:HiddenField ID="hdf_file_path" runat="server" />
    </form>
</body>
</html>
<%
    If isGenDownload Then
            response.flush()
           genDownloadFile()
    End If
%>