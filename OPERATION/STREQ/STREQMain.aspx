<%@ Page Language="VB" AutoEventWireup="false" CodeFile="STREQMain.aspx.vb" Inherits="OPERATION_STQ_STREQMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Stock Transfer</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">

function maskKey(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
    return false;
}

function maskDate(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
    return false;
}

function maskTel(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
    return false;
}

function ItemLookUp(STORER_CODE, WH_CODE) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2, qtyList|CC_ITEM_QTY#txt_item_qty");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}

function selectedToWH(rowIdx) {
    document.myform.selectedrowIndex.value = rowIdx;
    document.myform.moduleAction.value = "SELECTTOWH";
    document.myform.submit();
}

function LocLookUp(seq, rowIdx, lb_id, hd_id) {
    //if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp" + seq, "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");
        
        if (seq == 1) {
            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TQ_WH_FR.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TQ_WH_TO.value);
            
        }

        //setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedToWH('" + rowIdx + "')");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp" + seq;
        document.hiddenForm.submit();
    //}
}

function ItemLocLookUp(STORER_CODE, ITEM_CODE, PACK_KEY, lb_id, hd_id, pallet_id, hpallet_id, qty_id,lbatch_no, batch_no, tPallet_id, tbatch_no) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLocLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,left=5,top=15");

        
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_LOC");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "ic", ITEM_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pKey", PACK_KEY);
        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TQ_WH_FR.value);

        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + qty_id + "|3, " + lbatch_no + "|4|L, " + batch_no + "|4");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLocLookUp";
        document.hiddenForm.submit();
    }
}

function goToAttach(doc_type, doc_code, p_editmode) {

    //alert("Alert");
    removeAllElementFromForm(document.hiddenForm);
    window.open("", "NewAttachment", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=960,height=600,left=80,top=80");

    setInterfaceDataToForm(document.hiddenForm, "UPFL_DOC_NO", doc_code);
    setInterfaceDataToForm(document.hiddenForm, "DOC_TYPE", doc_type);
    setInterfaceDataToForm(document.hiddenForm, "PARENT_EDIT_MODE", p_editmode);

    document.hiddenForm.action = "../../ATTACH/ATTACH_MAIN.ASPX";
    document.hiddenForm.target = "NewAttachment";
    document.hiddenForm.submit();
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />    
    <asp:HiddenField ID="seqList" runat ="server" />
     <asp:HiddenField ID="qtyList" runat ="server" />
    <asp:HiddenField ID="selectedrowIndex" runat ="server" />    
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="3" class="TITLE" align="left">
                    <asp:label id="lheader" runat ="server" Text="Request for Transfer Maintenance" />
                </td>
                <td colspan="1" class="TITLE" align="right">
                    <input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" />
                </td>
            </tr>            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                    <%If Session("pagemode") <> "N" Then%>     
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" Visible="false" />
                        <asp:Button ID="btnClose" Text="Close Request Order" CssClass="all_button" runat="server" />
                        <asp:ConfirmButtonExtender runat="server" ID="cfm1" TargetControlID="btnClose" ConfirmText="This Inter-Store Transfer will be close.&#10;Confirm to Proceed?" />
                    <% End If%>
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="TQ_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="TQ_STATUS" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_TQ_EDI_RFT_NO" runat="server" /></font>
                </td>
                <td id="Td1" colspan="3" runat="server">
                    <font size="2">
                         <asp:TextBox ID="TQ_EDI_RFT_NO" runat="server" MaxLength="10"></asp:TextBox>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="TQ_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="TQ_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="TQ_BY" runat="server" MaxLength="20"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_BATCH_NO" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                      <asp:TextBox ID="TQ_BATCH_NO" MaxLength="20" runat="server" /></font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_REF_NO" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="TQ_REF_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_WH_FR" runat="server" /></font>
                </td>
                <td width="26%">
                    <font size="2">
                         <asp:DropDownList ID="TQ_WH_FR" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_WH_TO" runat="server" /></font>
                </td>
                <td width="26%">
                    <font size="2">
                         <asp:DropDownList ID="TQ_WH_TO" runat="server" />
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_CLOSED_DATE" runat="server" text="Closed Date" /></font>
                </td>
                <td>
                    <font size="2">
                      <asp:TextBox ID="TQ_CLOSED_DATE" MaxLength="10" runat="server" />
                      <asp:ImageButton ID="btnCdate" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="TQ_CLOSED_DATE" PopupButtonID="btnCdate" Format="dd/MM/yyyy" /></font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_DATA_FR" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:Label runat="server" ID="TQ_DATA_FR" />
                    </font>
                </td>                
            </tr>
            <tr runat="server" id="to_loc_tr" visible="false">
                <td class="LabelTD" nowrap width="15%">
                    
                </td>
                <td id="Td2" width="26%" runat="server">
                    
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_TO_LOC" runat="server" /></font>
                </td>
                <td id="Td3" width="26%" runat="server">
                    <font size="2">
                        <asp:Label ID="dsp_TQ_TO_LOC" runat="server" /></font>
                        <asp:HiddenField ID="TQ_TO_LOC" runat="server" />
                        <asp:Image ID="Image_Loc_LookUp_TQ_TO_LOC" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TQ_REM" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="TQ_REM" runat="server" Height="117px" Width="523px" 
                        TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td colspan="4" align="left" width="100%">
                    <table border="1" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_cb" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cb" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lub" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lub" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_cd" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_cd" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_sys_lud" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="2">
                                    <asp:Label ID="sys_lud" runat="server" />
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            </table>
            <table  border="0" cellspacing="1" cellpadding="1" align="center" style="width:100%">
              <tr>
                <td class="TITLE" colspan="4">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b><asp:Label ID="lbl_ImageHd" runat="server" /></b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:UpdatePanel runat="server" ID="LOOKUPUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" />    
                        </ContentTemplate>
                    </asp:UpdatePanel>                    
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="TQD_seq" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                             <asp:BoundField DataField="TQD_seq" HeaderText="No." >
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" />
                            </asp:BoundField>
                               <asp:TemplateField HeaderText="Stock No.">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%--<asp:TextBox ID="itm_name" runat="server" Font-Size="11px" Width="110px" MaxLength="100"></asp:TextBox>--%>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="itm_sku_no" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                                    <asp:HiddenField runat="server" id="itm_code" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="pack_key" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="itm_name" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>    
                            <asp:TemplateField HeaderText="Batch No." Visible="false">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" id="TQD_batch_no_fr" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Pallet No." Visible="false">
                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                <ItemTemplate>
                                    <asp:textbox ID="TQD_pallet_no_fr" runat="server" MaxLength="80" Font-Size="11px" Width="40px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="TQD_UOM" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <HeaderStyle HorizontalAlign="right" Width="65px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="TQD_qty" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="TQD_POST_QTY" HeaderText="Transferred Qty." ItemStyle-HorizontalAlign="Right" ItemStyle-Width="50px" ItemStyle-Font-Size="11px" />
                            <asp:TemplateField HeaderText="Serial No." Visible="false">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Width="50px" Font-size="11px" id="TQD_SERIAL" MaxLength="80" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="TQD_UOM2" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty2">
                                <HeaderStyle HorizontalAlign="right"  Width="65px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="TQD_qty2" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="TQD_rem" runat="server" Width="150px" Font-Size="10px" MaxLength="200"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                        </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
