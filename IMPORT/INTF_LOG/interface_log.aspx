<%@ Page Language="VB" AutoEventWireup="false" CodeFile="interface_log.aspx.vb" Inherits="IMPORT_INTF_LOG_intferface_log" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
<title>Interface Log</title>
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script type="text/javascript" language="javascript">

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="myform" runat="server">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td style="text-align:center; background-color:White">
            <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="4" class="TITLE" style="text-align:left">Interface Log</td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IMP_BATCH_ID" Text="Batch ID" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="IMP_BATCH_ID" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IMP_FILE_NAME" Text="File Name" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="IMP_FILE_NAME" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_IMP_TYPE" Text="Interface Type" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_IMP_TYPE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_STATUS" Text="Status" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_START_DT" Text="Start Date" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_START_DT" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_END_DT" Text="End Date" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_END_DT" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_ERRCODE" Text="Error Code" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_ERRCODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ITF_ERRDESC" Text="Error Description" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="ITF_ERRDESC" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="TITLE" style="text-align:left">
                    <input id="btnBack" runat="server" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=ADM_IFL'" class="all_button" />
                </td>
            </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td style="background-color:White">&nbsp;</td>
    </tr>
    <tr id="trPO" runat="server" visible="false" style="text-align:center">
        <td>
            <asp:GridView ID="gvPO" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left"
                HeaderStyle-CssClass="LabelTD" HeaderStyle-Font-Size="12px" >
            <Columns>
                <asp:BoundField DataField="IMP_ROWNUM" HeaderText="Row No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ALLOW_IMP_YN" HeaderText="Allow Import Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMPORTED_YN" HeaderText="Imported Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMP_CODE" HeaderText="Error Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_REMARKS" HeaderText="Error Desc." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_PO_NO" HeaderText="PO No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_REV_NO" HeaderText="Revision No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_VEND_NO" HeaderText="Vendor Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_ORD_DATE" HeaderText="Order Date" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_SHIP_ADDR" HeaderText="Ship To" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_OLD_VENDOR" HeaderText="Old Vender Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="POH_OLD_ORD_DATE" HeaderText="Old Order Date" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ACTION" HeaderText="RO Action" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="RO_CODE" HeaderText="RO Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="RO_WH_CODE" HeaderText="WH Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
            </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr id="trIR" runat="server" visible="false" style="text-align:center">
        <td>
            <asp:GridView ID="gvIR" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left"
                HeaderStyle-CssClass="LabelTD" HeaderStyle-Font-Size="12px" >
            <Columns>
                <asp:BoundField DataField="IMP_ROWNUM" HeaderText="Row No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ALLOW_IMP_YN" HeaderText="Allow Import Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMPORTED_YN" HeaderText="Imported Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMP_CODE" HeaderText="Error Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_REMARKS" HeaderText="Error Desc." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IRH_IR_NO" HeaderText="SIR No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IRH_DELIV_TO" HeaderText="Deliver To" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IRH_WH_NOS" HeaderText="WH ISS Ticket" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ACTION" HeaderText="CO Action" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="CO_CODE" HeaderText="CO Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                 <asp:BoundField DataField="IMP_ACTION2" HeaderText="DO Action" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="DO_CODE" HeaderText="DO Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
            </Columns>
            </asp:GridView>
        </td>
    </tr>
    <tr id="trTR" runat="server" visible="false" style="text-align:center">
        <td>
            <asp:GridView ID="gvTR" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left"
                HeaderStyle-CssClass="LabelTD" HeaderStyle-Font-Size="12px" >
            <Columns>
                <asp:BoundField DataField="IMP_ROWNUM" HeaderText="Row No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ALLOW_IMP_YN" HeaderText="Allow Import Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMPORTED_YN" HeaderText="Imported Y/N" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_IMP_CODE" HeaderText="Error Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_REMARKS" HeaderText="Error Desc." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_TR_NO" HeaderText="RFT No." >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_FROM_LOC" HeaderText="Source Warehouse" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_TO_LOC" HeaderText="Destination Warehouse" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_ORIG_BY" HeaderText="Originator" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_ORIG_DATE" HeaderText="Originated Date" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_ORIG_TIME" HeaderText="Originated Time" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_APPR_BY" HeaderText="Approver" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_APPR_DATE" HeaderText="Approved Date" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TRH_APPR_TIME" HeaderText="Approved Time" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="IMP_ACTION" HeaderText="TR Action" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
                <asp:BoundField DataField="TR_CODE" HeaderText="TR Code" >
                    <ControlStyle Width="30px"  />
                    <ItemStyle Font-Size="11px" />
                </asp:BoundField>
            </Columns>
            </asp:GridView>
        </td>
    </tr>
    </table>
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>