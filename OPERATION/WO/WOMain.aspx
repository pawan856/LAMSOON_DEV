<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WOMain.aspx.vb" Inherits="OPERATION_WO_Main" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Work Order</title>
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

function ItemLookUp(STORER_CODE) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}

function selectedItemLoc() {
    document.myform.moduleAction.value = "SELECTIMLOC";
    document.myform.submit();
}

function selectedToWH(rowIdx) {
    document.myform.selectedrowIndex.value = rowIdx;
    document.myform.moduleAction.value = "SELECTTOWH";
    document.myform.submit();
}

function LocLookUp(seq, rowIdx, lb_id, hd_id) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp" + seq, "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");
        
//        if (seq == 1) {
//            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.WO_WH_FR.value);
//        } else {
//            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.WO_WH_TO.value);
//            
//        }

        //setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedToWH('" + rowIdx + "')");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp" + seq;
        document.hiddenForm.submit();
    }
}

function ItemLocLookUp(STORER_CODE, ITEM_CODE, PACK_KEY, lb_id, hd_id, pallet_id, hpallet_id, qty_id,lbatch_no, batch_no, tPallet_id, tbatch_no) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLocLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=768,left=5,top=5");

        
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_LOC");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "ic", ITEM_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pKey", PACK_KEY);

        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + qty_id + "|3, " + lbatch_no + "|4|L, " + batch_no + "|4");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLocLookUp";
        document.hiddenForm.submit();
    }
}

function limitText(limitField, limitNum) {
    if (limitField.value.length > limitNum) {
        limitField.value = limitField.value.substring(0, limitNum);
    }
}

function ItemLocBalLookUp(STORER_CODE) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }
    
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLocLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_ILOC_BAL");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItemLoc()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2, seqList");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLocLookUp";
        document.hiddenForm.submit();
    }
}


function countKit(txtBox) {
    var bagQty = txtBox.value;
    var totalQty, kitQty;

    if (!isNaN(bagQty)) {
        
        var DataGridObj = document.getElementById('GridView1');
        if (DataGridObj.rows.length > 0) {
            for (var iRow = 1; iRow <= DataGridObj.rows.length - 1; iRow++) {
                kitQty = DataGridObj.rows[iRow].cells[1].getElementsByTagName("input")[0].value;
                if (DataGridObj.rows[iRow].cells[1].getElementsByTagName("select")[0].value == 'OUT' && !isNaN(kitQty)) {
                    totalQty = kitQty * bagQty;
                    DataGridObj.rows[iRow].cells[11].getElementsByTagName("input")[0].value = totalQty; 
                }
            }
        }
    }
}
</script>
<style type="text/css">
    .BlackCircle
    {
                width:30px;
                height:30px;
                border-radius:15px;
                background:#000;
                line-height:30px;
                text-align:center;
    }    
    
    .RedCircle
    {
                width:30px;
                height:30px;
                border-radius:15px;
                background:#F00;
                line-height:30px;
                text-align:center;
    }    
    
    .GreenCircle
    {
                width:30px;
                height:30px;
                border-radius:15px;
                background:#30C23A;
                line-height:30px;
                text-align:center;
    }
    .YellowCircle
    {
                width:30px;
                height:30px;
                border-radius:15px;
                background:#F0E80A;                
                line-height:30px;
                text-align:center;
    }
    
    
        
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
   <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />    
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />    
    <asp:HiddenField ID="seqList" runat ="server" />
    <asp:HiddenField ID="selectedrowIndex" runat ="server" />    
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="4" class="TITLE">
                <asp:label ID="lheader" runat ="server" />
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                </td>
            </tr>
            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                    <%If Session("pagemode") <> "N" Then%>    
                        <asp:button runat="server" ID="btnGenSR" text="Generate Stock Relocation" CssClass="all_button" />         
                        <asp:button runat="server" ID="btnCopy" text="Copy New WO" CssClass="all_button" />
                        <asp:ConfirmButtonExtender runat="server" ID="btnCopyCFM" TargetControlID="btnCopy" ConfirmText="A new WO will be created. Are you sure to proceed?" />
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" />
                    <% End If%>
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WO_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="WO_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WO_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="WO_STATUS" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>
                 <td class="LabelTD" nowrap>
                     <font size="2">
                        <asp:Label ID="lbl_WO_TYPE" runat="server" text="WO Type"/></font>:
                </td>
                <td>
                    <font size="2">
                         <asp:DropDownList ID="WO_TYPE" runat="server" />                           
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WO_DATE" runat="server" /></font>
                </td>
                <td colspan="1">
                    <asp:TextBox ID="WO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="WO_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WO_TR_CODE" runat="server" /></font>
                </td>
                <td colspan="1">
                    <asp:TextBox ID="WO_TR_CODE" runat="server"></asp:TextBox>&nbsp;<asp:LinkButton ruant="server" ID="gotoSTKREC" text="GoTo" runat="server" /> 
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WO_REM" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="WO_REM" runat="server" Height="117px" Width="523px" 
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
            <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 100%">
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
                <td colspan="3" class="menuTD">
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" />
                    <asp:Button runat="server" ID="selectLocBtn" Text="Select Item From Balanace" CssClass="all_button" />       
                    
                </td>
                <td class="menuTD">
                <asp:UpdatePanel runat="server" RenderMode="Inline" id="chkBalUDP">
                        <ContentTemplate>
                          <asp:button runat="server" ID="btnChkBal" CssClass="all_button" Text="Check Stock Balance" Visible="false" />
                        </ContentTemplate>
                    </asp:UpdatePanel>           
                </td>
            </tr>
            <tr>
                <td colspan="4"><asp:UpdatePanel runat="server" ID="GVUDP" RenderMode="Inline">
                <ContentTemplate>
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="WOD_seq" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                             <asp:BoundField DataField="WOD_seq" HeaderText="No." >
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="In / Out">
                                <ControlStyle Width="150px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList runat="server" id="WOD_TYPE">
                                        <asp:ListItem Value="IN" text="In" />
                                        <asp:ListItem Value="OUT" Text="Out" />
                                    </asp:DropDownList>
                                    <asp:HiddenField runat="server" ID="WOD_KIT_QTY" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Internal Item Code">
                                <ControlStyle Width="150px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%--<asp:TextBox ID="itm_code" runat="server" Font-Size="11px" Width="150px" MaxLength="20"></asp:TextBox>--%>
                                    <asp:Label runat="server" Font-size="11px" id="itm_code" />
                                    <asp:HiddenField runat="server" ID="WOD_seq" />
                                    <asp:HiddenField runat="server" ID="itm_sku_no" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:BoundField runat="server" HeaderText="Stock No." DataField="itm_sku_no" ItemStyle-Font-Size="11px" />
                             <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%--<asp:TextBox ID="pack_key" runat="server" Font-Size="11px" Width="60px" MaxLength="50"></asp:TextBox>--%>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="pack_key" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%--<asp:TextBox ID="itm_name" runat="server" Font-Size="11px" Width="110px" MaxLength="100"></asp:TextBox>--%>
                                    <asp:Label runat="server" Font-size="11px" id="itm_name" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Textbox id="wod_batch_no" runat="server" MaxLength="20" Font-Size="11px" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Serial No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Textbox id="WOD_SERIAL_NO" runat="server" MaxLength="20" Font-Size="11px" />
                                    <asp:hiddenfield ID="WOD_pallet_no" runat="server" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Drum ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Textbox id="WOD_DRUM_ID" runat="server" MaxLength="20" Font-Size="11px" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Drum LV">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Textbox id="WOD_DRUM_LV" runat="server" MaxLength="20" Font-Size="11px" width="50px" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Loc.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                     <asp:DropDownList ID="wod_loc" runat="server" width="100px" Font-Size="11px"></asp:DropDownList>

                                        <%--<table border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_wod_loc" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td>
                                                    <asp:Image ID="Image_Loc_LookUp_To" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="wod_loc" runat="server" />--%>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="WOD_qty" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="itm_UOM" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty2">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="WOD_qty2" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="itm_UOM2" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="WOD_rem" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                        <asp:Button ID="btnRFID" runat="server" Font-Size="11px"
                                        CommandName="RFID" Text="RFID No." CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>                                
                                <HeaderStyle Width="80px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="No. of RFID Tag" HeaderStyle-Width="60px" ItemStyle-HorizontalAlign="Right" >
                                <ItemTemplate>
                                      <asp:Label runat="server" ID="RFID_COUNT" Font-Size="11px" />                                    
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
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
                    </ContentTemplate>
                </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td class="TITLE" colspan="4">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b><asp:Label ID="hd_res" runat="server" Text="Work Order Resources" /></b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="btnAddRes" runat="server" Text="Add Resources" CssClass="all_button" UseSubmitBehavior ="false" />
                </td>
            </tr>
            <tr>
                <td colspan="4">                
                    <asp:GridView ID="GridView2" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="WO_RES_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                             <asp:BoundField DataField="WO_RES_SEQ" HeaderText="No." >
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Type">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                   <asp:DropDownList runat="server" ID="WO_RES_TYPE">
                                    <asp:ListItem Value="H" Text="HR" />
                                    <asp:ListItem Value="C" Text="COG" />
                                    <asp:ListItem Value="O" Text="OTHERS" />
                                   </asp:DropDownList> 
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox runat="server" ID="WO_RES_DESC" MaxLength="200" width="200px" class="REQUIRED" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Amount">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox runat="server" ID="WO_RES_AMT" MaxLength="12" onkeypress="return maskKey(event);" class="REQUIRED" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox runat="server" ID="WO_RES_REM" MaxLength="500" Width="250px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="List">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox runat="server" ID="WO_RES_LIST" MaxLength="500" Width="250px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDeleteR" name="btnDeleteR" runat="server" Height="22px" Font-Size="11px"
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
                <td colspan="3" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
                <td align="right" class="menuTD"><asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" /></td>
                                        
            </tr>
        </table>
    </div>

    <asp:UpdatePanel runat="server" ID="udp1" RenderMode="Inline">
        <ContentTemplate>
                <asp:Panel ID="pnlRFID" runat="server" CssClass="modalPopup" style="display: none" ><!---->
                    <table width="500px">
                       <tr>
                            <td class="TITLE" colspan="2">
                                <asp:Panel runat="Server" ID="PanelDrag1" Style="cursor: move;">
                                <font size="2">RFID No List:</font>
                                </asp:Panel>
                            </td>
                       </tr>
                       <tr>
                            <td>
                               <asp:Button runat="server" ID="btnAddRFID" CssClass="all_button" Text="Add RFID Tag" Visible="false" />
                            </td>
                            <td align="right">
                                <asp:Button runat="server" ID="btnPnlClose" Text="Close" CssClass="all_button" />
                            </td>
                       </tr>
                       <tr>
                          <td colspan="2">
                               <asp:GridView ID="GVRFID" runat="server" Height="10px" Width="400px" Font-Names="Arial"
                                    Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                    CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                                    <RowStyle CssClass="GV" />
                                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                                    <Columns>
                                         <asp:TemplateField HeaderText="RFID Code">
                                            <HeaderStyle HorizontalAlign="Left" Width="150px" />
                                            <ItemTemplate>
                                                <asp:TextBox runat="server" ID="WORF_RFID" MaxLength="24" Font-Size="11px" Width="98%" />
                                                <asp:HiddenField runat="server" ID="WOD_SEQ" />
                                                <asp:HiddenField runat="server" ID="WORF_SEQ" />
                                            </ItemTemplate>                                            
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Status" HeaderStyle-Width="60px">
                                            <HeaderStyle HorizontalAlign="center" />
                                            <ItemTemplate>
                                                <asp:panel runat="server"  ID="WORF_STATUS" HorizontalAlign="Center">
                                                    <asp:Label runat="server" ID="lbl_status" Font-Size="9px" text="NEW" ForeColor="White" />
                                                </asp:panel> 
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderStyle-Width="60px">
                                            <HeaderStyle HorizontalAlign="center" />
                                            <ItemTemplate>
                                                <asp:Button runat="server" ID="btnOut" Text="Out" CssClass="all_button" CommandName="OUT" />
                                                <asp:ConfirmButtonExtender runat="server" ID="outCFM" ConfirmText="Are you sure to change stauts as OUT?" TargetControlID="btnOut" />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField ControlStyle-Width="50px">
                                            <ItemTemplate>
                                                <asp:Button ID="btnDelete" name="btnDeleteR" runat="server" Height="22px" Font-Size="11px"
                                                    CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                            </ItemTemplate>
                                            <ControlStyle Width="50px"></ControlStyle>
                                            <HeaderStyle Width="50px" />
                                        </asp:TemplateField>
                                    </Columns>
                               </asp:GridView> 
                          </td>
                       </tr>                      
                    </table>
                </asp:Panel> 
                <asp:HiddenField runat="server" ID="dummy" />
                <asp:ModalPopupExtender ID="pnlRFID_ModalPopupExtender" runat="server"
                    DynamicServicePath="" 
                    Enabled="True" 
                    TargetControlID="dummy" 
                    PopupControlID="pnlRFID"
                    BackgroundCssClass="modalBackground"
                    DropShadow="true"         
                    CancelControlID="btnPnlClose"
                    Y="50"
                    PopupDragHandleControlID="PanelDrag1" RepositionMode="None">
                </asp:ModalPopupExtender>        
        </ContentTemplate>
    </asp:UpdatePanel>


    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
