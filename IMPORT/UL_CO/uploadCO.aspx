<%@ Page Language="VB" AutoEventWireup="false" CodeFile="uploadCO.aspx.vb" Inherits="EXCEL_uploadCO" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
<title>Import CO</title>
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<div runat="server">
<script type="text/javascript" language="javascript">
//    function showLogList() {
//        removeAllElementFromForm(document.hiddenForm);
//        setInterfaceDataToForm(document.hiddenForm, "user_id", document.uploadform.user_id.value);
//        document.hiddenForm.action = "./APLogListing.aspx";
//        document.hiddenForm.target = "_self";
//        document.hiddenForm.submit();
//    }s
    function showLink(alink) {

        document.getElementById('link_span').innerHTML = "<a href='" + alink + "'>Click to Created CO</a>";
    }
</script>
</div>
</head>
<body id="b1">
    <form name="uploadform" id="uploadform" runat="server" enctype="multipart/form-data">
    <asp:ToolkitScriptManager runat="server" id="SC1" />
    <div>
    <table border="0" cellspacing="1" cellpadding="2" width="100%">
	<tr style="background-color:#006699">
		<td class="TITLE" colspan="2"><b><asp:label runat="server" ID="lblHD" text="Upload CO" /></b></td>
	</tr>
	<tr>
		<td class="TITLE" colspan="2" valign="middle">
		    <asp:Button runat="server" ID="btnReset" Text="Reset" CssClass="all_button"  OnClientClick="Javascript:window.location='.\uploadCO.aspx'; return false();" />            
           &nbsp;&nbsp;<asp:Button runat="server" ID="btnDL" Text="Download Import Excel Template" CssClass="all_button" />
		</td>
	</tr>
	<tr>
		<td class="LabelTD" valign="top" width="15%"><asp:label runat="server" ID="lbl_OUTPUT" Text="Output" /> :</td>
		<td>
		    <asp:Label ID="outputSpan" runat="server">---</asp:Label>
		</td>
	</tr>
	<tr>
		<td class="LabelTD" valign="top" width="15%"><asp:label runat="server" ID="lbl_status" Text="Status" /> :</td>
		<td>
		    <asp:Label ID="statusSpan" runat="server">Please upload your CO EXCEL Data file.</asp:Label>
		</td>
	</tr>
    <tr>
       <td class="LabelTD" valign="top" width="15%"><asp:label runat="server" ID="lbl_storer_code" Text="Storer" /> :</td>
		<td>
          <asp:UpdatePanel runat="server" ID="UDP1" RenderMode="Inline">
            <ContentTemplate>
                <asp:DropDownList runat="server" id="STORER_CODE" class="REQUIRED" AutoPostBack="true" />         
            </ContentTemplate>
          </asp:UpdatePanel>		   
		</td>
    </tr>
    <tr runat="server" visible="false" >
        <td class="LabelTD" valign="top" width="15%">Customer:</td>
        <td>
         <asp:UpdatePanel runat="server" ID="UpdatePanel1" RenderMode="Inline">
            <ContentTemplate>
                <asp:DropDownList runat="server" id="CUS_CODE">
                    <asp:ListItem Value="" Text="SELECT" />
                </asp:DropDownList>
            </ContentTemplate> 
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="STORER_CODE" />
            </Triggers>
         </asp:UpdatePanel>             
        </td>
    </tr>
	<div id="TR_FILE" runat="server">
	<tr>
		<td class="LabelTD" valign="top" width="15%"><asp:label runat="server" ID="lbl_datafile" Text="Data File" /> :</td>
		<td>
            <asp:FileUpload ID="FileUpload1" runat="server" style="vertical-align:middle" />&nbsp;&nbsp;
            <asp:Button runat="server" ID="btnSubmit" Text="Submit"  CssClass="all_button" />
		</td>
	</tr>
    </div>
    <tr>
        <td colspan="2" class="LabelTD" align="left">
            <asp:label runat="server" ID="lbl_result" Text="Upload Result" />:
        </td>
    </tr>
    <tr>
        <td colspan="2">
            <asp:textbox runat="server" ReadOnly TextMode="MultiLine" Width="80%" ID="up_result" Rows="8" />
        </td>
    </tr>
	<tr>
	    <td>&nbsp;</td>
		<td>
		     <span id="link_span">
	          
	         </span>
	    </td>
	</tr>
	<tr>
		<td class="TITLE" colspan="2">
            &nbsp;
		</td>
	</tr>    
    </table>
    </div>
    <asp:HiddenField runat="server" ID="NOT_UPDATED_ROW" />
    <asp:HiddenField ID="user_id" runat="server" />
    </form>
    
    <form name="hiddenForm" method="post"></form>

    
    <iframe src="../blank.html" name="ifrm" id="ifrm" width="0" height="0"></iframe>
</body>
</html>
<%
    If isSubmit Then
        uploadFile()
    End If
%>