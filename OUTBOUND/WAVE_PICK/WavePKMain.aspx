<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WavePKMain.aspx.vb" Inherits="OUTBOUND_WAVE_PICK" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Stock Adjustment</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<div runat="server" id="SCRIPTDIV">
<script language="javascript">

function limitText(limitField, limitNum) {
        if (limitField.value.length > limitNum) {
            limitField.value = limitField.value.substring(0, limitNum);
        }
    }

function LocLookUp(lb_id, hd_id, lqty_id, hqty_id, ic, pk) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");

        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.AD_WH.value);
        setInterfaceDataToForm(document.hiddenForm, "sc", document.myform.STORER_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "ic", ic);
        setInterfaceDataToForm(document.hiddenForm, "pack", pk);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + lqty_id + "|QL, " + hqty_id + "|Q, " + hd_id);
        
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function calQty(rQty, loqid, lvqid, oqid, vqid) {
    var vQty;
    var oQty = eval("document.getElementById('" + oqid + "')").value;

    if (rQty == '') rQty = oQty;

    if (parseInt(oQty) > parseInt(rQty)) {
        vQty = (parseInt(oQty) - parseInt(rQty));
        eval("document.getElementById('" + lvqid + "')").innerHTML = "-" + vQty;
        eval("document.getElementById('" + vqid + "')").value = "-" + vQty;
    } else if (parseInt(oQty) < parseInt(rQty)) {
        vQty = (parseInt(rQty) - parseInt(oQty));
        eval("document.getElementById('" + lvqid + "')").innerHTML = vQty;
        eval("document.getElementById('" + vqid + "')").value = vQty;
    } else if (parseInt(oQty) == parseInt(rQty)) {
    eval("document.getElementById('" + lvqid + "')").innerHTML = 0;
        eval("document.getElementById('" + vqid + "')").value = 0;
    }
}

function ItemLocLookUp(STORER_CODE, ITEM_CODE, PACK_KEY, lb_id, hd_id, pallet_id, hpallet_id, lQty_id, qty_id, lbatch_no, batch_no, lexp_date, exp_date, lmanu_date, manu_date) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLocLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=500,left=5,top=15");

        
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_LOC");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "ic", ITEM_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pKey", PACK_KEY);
        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.AD_WH.value);

        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + lQty_id + "|3|L, " + qty_id + "|3, " + lbatch_no + "|4|L, " + batch_no + "|4, " + lexp_date + "|5|L, " + exp_date + "|5, " + lmanu_date + "|6|L, " + manu_date + "|6 ");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLocLookUp";
        document.hiddenForm.submit();
    }
}

function DOLookUp() {
    if (document.getElementById('STORER_CODE').value == '') {
        alert('Please select the Storer first!');
        return false;
    }

    if (document.getElementById('WP_WH_CODE').value == '') {
        alert('Please select the Warehouse first!');
        return false;
    } 
        else
    {
        //document.getElementById('WP_WH_CODE').disabled = true;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "DOLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=750,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_DOWP");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedDO()");
        setInterfaceDataToForm(document.hiddenForm, "wh", document.getElementById('WP_WH_CODE').value);
        setInterfaceDataToForm(document.hiddenForm, "sc", document.getElementById('STORER_CODE').value);
        //setInterfaceDataToForm(document.hiddenForm, "pItemList", GR_DOC_NO + "|1");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "DOList|1");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "DOLookUp";
        document.hiddenForm.submit();
    }
}

function LocLookUp(itm_code, pack_key, pallet_no, batch_no, sku_no,do_code,co_code) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");

        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "IMP_CODE", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById('STORER_CODE').value);
        setInterfaceDataToForm(document.hiddenForm, "ITM_CODE", itm_code);
        setInterfaceDataToForm(document.hiddenForm, "ITM_SKU_NO", sku_no);
        setInterfaceDataToForm(document.hiddenForm, "PACK_KEY", pack_key);
        setInterfaceDataToForm(document.hiddenForm, "PALLET_NO", pallet_no);
        //setInterfaceDataToForm(document.hiddenForm, "ref_no", opener.myform.DO_CUS_REF_NO.value);
        setInterfaceDataToForm(document.hiddenForm, "BATCH_NO", batch_no);

        setInterfaceDataToForm(document.hiddenForm, "CO_CODE", co_code);
        setInterfaceDataToForm(document.hiddenForm, "DO_CODE", do_code);

        document.hiddenForm.action = "./PickListLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function selectedDO() {
    document.myform.moduleAction.value = "SELECTDO";
    document.myform.submit();
}

function refreshGV() {
    document.myform.moduleAction.value = "RELOADPL";
    document.myform.submit();
}

function PrintPickList() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintPickList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1150,height=576,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);
        }
        
        setInterfaceDataToForm(document.hiddenForm, "WP_NO", document.getElementById('WP_NO').innerHTML);
        document.hiddenForm.action = "picklist_print.aspx";
        document.hiddenForm.target = "doPrintPickList";
        document.hiddenForm.submit();       
}

function OpenDO(do_code) {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1150,height=576,left=5,top=15");
               
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById('STORER_CODE').value);
        
        setInterfaceDataToForm(document.hiddenForm, "DO_CODE", do_code);
        setInterfaceDataToForm(document.hiddenForm, "newWin", 'Y');
        document.hiddenForm.action = "../DO/DOMain.aspx";
        document.hiddenForm.target = "doMain";
        document.hiddenForm.submit();       
}
</script>
</div>
<style type="text/css">

</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="DOList" runat ="server" />
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />
    <asp:HiddenField ID="selectedDO" runat ="server" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:label ID="lheader" runat ="server" />
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
                        <%If Viewstate("pagemode") <> "N" Then%>             
                            <asp:Button runat="server" ID="btnRelease" Text="Release" CssClass="all_button" />                        
                            <asp:Button runat="server" ID="btnReloadPicking" Text="Release" CssClass="all_button" />
                        <% End If%>
                        </td>
                    </tr>
                   </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WP_NO" runat="server" Text="Wave Picking No.:" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="WP_NO" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WP_STATUS" runat="server" Text="Status" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="WP_STATUS" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" Text="Storer:" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WP_WH" runat="server" Text="Warehouse:" /></font>
                </td>
                <td id="Td1" runat="server">
                    <font size="2">
                         <asp:DropDownList ID="WP_WH_CODE" runat="server" CssClass="REQUIRED" />
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WP_DATE" runat="server" Text="Date" />
                </td>                
                <td width="150px">
                    <asp:TextBox ID="WP_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" CssClass="REQUIRED"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="WP_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />  
                </td>
                <td>

                </td>
                <td>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CUS_CODE" runat="server" Text="Customer No.:" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="CUS_CODE" runat="server" MaxLength="80" Width="150px"></asp:TextBox>
                    </font>
                </td>     
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CUS_NAME" runat="server" Text="Customer Name:" /></font>
                </td>           
                <td>
                    <font size="2">
                        <asp:TextBox ID="CUS_NAME" runat="server" MaxLength="400" Width="150px"></asp:TextBox>
                    </font>
                </td>
            </tr>  
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DO_NO" runat="server" Text="Total DO:" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                       <asp:Label runat="server" id="DO_NO" />
                    </font>
                </td>     
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_PO_NO" runat="server" Text="Total PL:" /></font>
                </td>           
                <td>
                    <font size="2">
                        <asp:Label runat="server" id="PL_NO" />
                    </font>
                </td>
            </tr>          
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_WP_REMARK" runat="server" Text="Remarks:" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="WP_REMARK" runat="server" Height="117px" Width="523px" 
                        TextMode="MultiLine" onKeyDown="limitText(this,800);" onKeyUp="limitText(this,800);" onBlur="limitText(this,800);" />
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <hr style="width: 896px" />
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
            <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 100%">
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
                <td colspan="2" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" visible ="false" />
                    <asp:Button ID="btnSelDO" runat="server" Text="Select WIT" CssClass="all_button" OnClientClick="javascript:DOLookUp();return false;"/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button runat="server" ID="btnReset" Text="Reset PickList(Unused)" Visible="false" CssClass="all_button" />&nbsp;&nbsp;                    
                    <asp:Button runat="server" ID="btnSPLF" Text="Reset PL No." CssClass="all_button" />&nbsp;&nbsp;
                    <asp:Button runat="server" ID="btPL1" Text="Set PL No. to 1" CssClass="all_button" />&nbsp;&nbsp;
                    <asp:Button runat="server" ID="btnPrint" Text="Print PickList" CssClass="all_button" OnClientClick="javascript:PrintPickList();return false;" />
                </td>
                <td colspan="2" class="menuTD" style="border:none;">
                    <asp:label runat="server" ID="wtplNo" text="PL No.:" Font-Size="11px" Font-Bold="true" />
                    <asp:DropDownList runat="server" id="wtPLNO_SEL" Font-Size="11px">
                        <asp:ListItem Text="1" Value="1" /> 
                        <asp:ListItem Text="2" Value="2" /> 
                        <asp:ListItem Text="3" Value="3" /> 
                        <asp:ListItem Text="4" Value="4" /> 
                        <asp:ListItem Text="5" Value="5" /> 
                        <asp:ListItem Text="6" Value="6" /> 
                        <asp:ListItem Text="7" Value="7" /> 
                        <asp:ListItem Text="8" Value="8" /> 
                    </asp:DropDownList>
                    <asp:Button runat="server" ID="btnSPLW" Text="Split PL Grp by Weight" CssClass="all_button" Font-Size="11px" />&nbsp;&nbsp;
                    <asp:label runat="server" text="Sort: " id="sr1"  Font-Bold="true" Font-size="11px" />&nbsp;
                    <asp:DropDownList runat="server" ID="sort_type" Font-Size="11px">
                        <asp:ListItem Text="SKU" value="SKU" />
                        <asp:listitem Text="WIT No." value="DO" />
                    </asp:DropDownList>&nbsp;&nbsp;
                    <asp:label runat="server" text="PL No." id="sr2"  Font-Bold="true" Font-size="11px" />&nbsp;
                    <asp:DropDownList runat="server" ID="sort_pl_list_no" Font-Size="11px">
                        <asp:ListItem Text="ALL" Value="ALL" />
                        <asp:ListItem Text="1" Value="1" /> 
                        <asp:ListItem Text="2" Value="2" /> 
                        <asp:ListItem Text="3" Value="3" /> 
                        <asp:ListItem Text="4" Value="4" /> 
                        <asp:ListItem Text="5" Value="5" /> 
                        <asp:ListItem Text="6" Value="6" /> 
                        <asp:ListItem Text="7" Value="7" /> 
                        <asp:ListItem Text="8" Value="8" /> 
                    </asp:DropDownList>
                    <asp:Button runat="server" id="btnSortPL" text="Sort SUB PL" CssClass="all_button" Font-Size="11px" />
                    <asp:RadioButtonList runat="server" ID="pl_sort_option" AutoPostBack="true"
                        RepeatDirection="Horizontal" style="background-color:transparent;" 
                        RepeatLayout="Flow" Visible="false">
                        <asp:ListItem Selected="True" Text="By DO" Value="DO" />
                        <asp:ListItem Text="By SKU" Value="SKU" />
                        <asp:ListItem Text="By LOC" Value="LOC" />
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                     <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                       Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                       EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left" >
                       <Columns>
                           <asp:TemplateField HeaderText="WSIR no.">
                               <ControlStyle Width="80px" />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                                   <asp:Hyperlink runat="server" ID="do_code" Font-Size="11px" />
                                 <%--  <asp:HiddenField runat="server" ID="do_co_code" />--%>
                                   <asp:hiddenfield ID="pld_seq" runat="server" />
                                   <asp:CheckBox ID="pld_is_loan" runat="server" Checked="false" Visible="false" />
                               </ItemTemplate>
                           </asp:TemplateField>    
                           <asp:TemplateField HeaderText="CO No.">
                               <ControlStyle Width="80px" />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                                   <asp:Label ID="DO_CO_CODE" runat="server" Font-Size="11px"></asp:Label>                                   
                                  <%-- <asp:Label ID="DO_EDI_SIR_NO" runat="server" Font-Size="11px"></asp:Label>--%>
                               </ItemTemplate>
                           </asp:TemplateField>                    
                           <asp:TemplateField HeaderText="Picked By">
                               <ControlStyle Width="80px" />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                                   <asp:TextBox ID="pld_picked_by" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Stock No.">
                               <ControlStyle Width="120px" />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                                   <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px"></asp:Label>
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Lot No.">
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemStyle Width="80px" />
                               <ItemTemplate>
                                   <asp:Label ID="pld_batch_no" runat="server" Font-Size="11px" maxlength="20" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Packy Key">
                               <ControlStyle Width="35px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:hiddenfield ID="pld_item_no" runat="server" />
                                   <asp:hiddenfield ID="pld_pallet_no" runat="server" />
                                   <asp:Label ID="pld_pack_key" runat="server" Font-Size="11px" style="text-align:center"></asp:Label>
                               </ItemTemplate>
                           </asp:TemplateField>                                   
                           <asp:TemplateField HeaderText="WSIR Qty">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:Label ID="pld_do_qty" runat="server" Font-Size="11px" style="text-align:right" />
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="FFI Qty">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:TextBox ID="pld_foi_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Picked Qty">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:TextBox ID="pld_item_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Stock Qty">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:Label ID="stock_qty" runat="server" Font-Size="11px" style="text-align:right" />
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Bal Qty">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:Label ID="iloc_bal_qty" runat="server" Font-Size="11px" style="text-align:right" />
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>

                             <asp:TemplateField HeaderText="Status">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:DropDownList ID="pld_status" runat="server" Font-Size="11px" style="text-align:right"></asp:DropDownList>
                                </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>

                           <asp:TemplateField HeaderText="Hold Qty / Total Bal" Visible="false">
                               <ControlStyle Width="80px" />
                               <HeaderStyle HorizontalAlign="Right" />
                               <ItemTemplate>
                                   <asp:Label ID="hold_qty_desc" runat="server" Font-Size="11px" style="text-align:right" />
                                   <asp:HiddenField ID="avail_qty" runat="server" />
                                   <asp:HiddenField ID="hold_qty" runat="server" />
                                   <asp:HiddenField ID="total_bal" runat="server" />
                                   <asp:HiddenField ID="total_foi_qty" runat="server" />
                                   <asp:HiddenField ID="total_item_qty" runat="server" />
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Right" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Location">
                               <ControlStyle />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                               <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                               <tr>
                                   <td align="left">
                                       <asp:Label ID="dsp_pld_loc" width="100px" runat="server" Font-Size="11px" />
                                       <asp:Label ID="dsp_pld_wh" width="70" runat="server" Font-Size="11px" style="text-align:center;display:none;" />
                                       <asp:HiddenField ID="pld_wh" runat="server" />
                                       <asp:HiddenField ID="pld_loc" runat="server" />
                                   </td>
                                   <td align="right">
                                       <asp:Image ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                                   </td>
                               </tr>
                               </table>
                               </ItemTemplate>
                           </asp:TemplateField>
                            <asp:TemplateField HeaderText="Weight Type">
                               <ControlStyle Width="60px" />
                               <HeaderStyle HorizontalAlign="Left" />
                               <ItemTemplate>
                                   <asp:Label ID="ITM_WEIGHT_TYPE" runat="server" Font-Size="11px" />
                               </ItemTemplate>
                               <ItemStyle HorizontalAlign="Left" />
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Floor">
                               <ControlStyle Width="35px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:Label ID="dsp_pld_floor" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                                   <asp:HiddenField ID="pld_floor" runat="server" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Area" Visible="false">
                               <ControlStyle Width="35px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:Label ID="dsp_pld_area" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                                   <asp:HiddenField ID="pld_area" runat="server" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="Rack" Visible="false">
                               <ControlStyle Width="35px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:Label ID="dsp_pld_rack" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                                   <asp:HiddenField ID="pld_rack" runat="server" />
                               </ItemTemplate>
                           </asp:TemplateField> 
                           <asp:TemplateField HeaderText="Bin" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:Label ID="dsp_pld_bin" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                                   <asp:HiddenField ID="pld_bin" runat="server" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 1" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL1" GroupName="PLD_PL_LIST_NO" value="1" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 2" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL2" GroupName="PLD_PL_LIST_NO" value="2" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 3" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL3" GroupName="PLD_PL_LIST_NO" value="3" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 4" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL4" GroupName="PLD_PL_LIST_NO" value="4" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 5" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL5" GroupName="PLD_PL_LIST_NO" value="5" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 6" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL6" GroupName="PLD_PL_LIST_NO" value="6" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 7" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL7" GroupName="PLD_PL_LIST_NO" value="7" />
                               </ItemTemplate>
                           </asp:TemplateField>
                           <asp:TemplateField HeaderText="PL 8" Visible="false">
                               <ControlStyle Width="25px" />
                               <HeaderStyle HorizontalAlign="Center" />
                               <ItemStyle HorizontalAlign="Center" />
                               <ItemTemplate>
                                   <asp:RadioButton runat="server" ID="PL8" GroupName="PLD_PL_LIST_NO" value="8" />
                               </ItemTemplate>
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
                       <AlternatingRowStyle CssClass="REQUIRED" />
                       <RowStyle CssClass="REQUIRED" />
                       <EmptyDataRowStyle CssClass="REQUIRED" />
                       <EditRowStyle CssClass="REQUIRED" />
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
                <td class="menuTD" align="right">
                    <table style=" border:none;">
                        <tr>
                            <td style="background-color:Transparent;">Status:</td>
                            <td width="20px" style="background-color:Yellow;">&nbsp;</td>
                            <td style="background-color:Transparent;">Assigned</td>
                            <td width="20px" style="background-color:Orange;">&nbsp;</td>
                            <td style="background-color:Transparent;">Picking</td>
                            <td width="20px" style="background-color:Green;">&nbsp;</td>
                            <td style="background-color:Transparent;">Short-Picked</td>
                            <td width="20px" style="background-color:Blue;">&nbsp;</td>
                            <td style="background-color:Transparent;">Picked</td>
                            <td width="20px" style="background-color:Red;">&nbsp;</td>
                            <td style="background-color:Transparent;">Cancelled</td>
                        </tr>
                    </table>
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

