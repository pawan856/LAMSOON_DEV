<%@ Page Language="VB" AutoEventWireup="false" CodeFile="uploadRO.aspx.vb" Inherits="EXCEL_uploadRO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
<title>Upload Non-Stock Item</title>
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script type="text/javascript" language="javascript">
//    function showLogList() {
//        removeAllElementFromForm(document.hiddenForm);
//        setInterfaceDataToForm(document.hiddenForm, "user_id", document.uploadform.user_id.value);
//        document.hiddenForm.action = "./APLogListing.aspx";
//        document.hiddenForm.target = "_self";
//        document.hiddenForm.submit();
//    }s
    function showLink(alink) {

        document.getElementById('link_span').innerHTML = "<a href='" + alink + "'>Click to view the Created PO</a>";
    }
</script>
</head>
<body id="b1">
    <form name="uploadform" id="uploadform" runat="server" enctype="multipart/form-data">
    <div>
    <table border="0" cellspacing="1" cellpadding="2" width="100%">
	<tr style="background-color:#006699">
		<td class="TITLE" colspan="2"><b>Upload Non-Stock Item</b></td>
	</tr>
	<tr>
		<td class="TITLE" colspan="2" valign="middle">
		    <asp:Button runat="server" ID="btnReset" Text="Reset" CssClass="all_button" />            
           &nbsp;&nbsp;<asp:Button runat="server" ID="btnDL" Text="Download Import Excel Template" CssClass="all_button" />
		</td>
	</tr>
	<tr>
		<td class="LabelTD" valign="top" width="15%">Output :</td>
		<td>
		    <asp:Label ID="outputSpan" runat="server">---</asp:Label>
		</td>
	</tr>
	<tr>
		<td class="LabelTD" valign="top" width="15%">Status :</td>
		<td>
		    <asp:Label ID="statusSpan" runat="server">Please upload your RO EXCEL Data file.</asp:Label>
		</td>
	</tr>
    <tr>
       <td class="LabelTD" valign="top" width="15%">Storer :</td>
		<td>
		   <asp:DropDownList runat="server" id="STORER_CODE" class="REQUIRED" />
		</td>
    </tr>
    <tr>
       <td class="LabelTD" valign="top" width="15%">Warehouse :</td>
		<td>
		   <asp:DropDownList runat="server" id="RO_WH_CODE" class="REQUIRED" />
		</td>
    </tr>
	<tr id="TR_FILE" runat="server">
		<td class="LabelTD" valign="top" width="15%">Data File :</td>
		<td>
            <asp:FileUpload ID="FileUpload1" runat="server" style="vertical-align:middle" />&nbsp;&nbsp;
            <asp:Button runat="server" ID="btnSubmit" Text="Submit"  CssClass="all_button" />
		</td>
	</tr>
    <tr id="EXP_TR" runat="server" style="display:none;">
		<td colspan="2">
            <asp:Button runat="server" ID="BTNOver" CssClass="all_button" Text="Overwrite PO with contract No." />
		</td>
	</tr>
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
    <asp:HiddenField ID="h_WIRO_BATCH_ID" runat="server" />
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