<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SRSRVMain.aspx.vb" Inherits="OPERATION_SRV_SRVMain" %>
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
<script language="javascript">

function ItemLookUp(STORER_CODE, WH) {
    if (STORER_CODE == '') {
        alert('Please select the Storer first!');
        return false;
    }
    
    if (WH == '') {
        alert('Please select the Warehouse first!');
        return false;
    }

    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_RO_MASTER");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}

function LocLookUp(lb_id, hd_id, cbm_id,area_id,type_id,cbmVal,perHidden,perLabel) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");

        var cbmText

        cbmText = document.getElementById(cbmVal).value

        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.LOBH_WH.value);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "fun_code", "OP_SRSRV");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + ", " + hd_id + ", " + cbm_id + ", " + area_id + ", " + type_id + ", " + cbmText + ", " + perHidden + ", " + perLabel);
          
        
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function maskKeyAndCheck(objEvent, loc) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    
    if (loc != null) {
        if (loc != '') {
            if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) {
                return true;
            } else {
                return false;
            }            
        } else {
            alert("You Must Select Location First!");
            return false;
        }

    } else {
            alert("You Must Select Location First!");
            return false;
    }
}

function calCBM(icbm, Bcbm, Barea, Btype, DSPper, HDPer) {
    var vBcbm, vBarea, vBtype
    var vPer

    vBcbm = document.getElementById(Bcbm).value;
    vBarea = document.getElementById(Barea).value;
    vBtype = document.getElementById(Btype).value;
    //alert(icbm);

    if (vBtype == 'AREA') {
        if (!isNaN(vBarea)) {
            if (vBarea != 0) {
                vPer = (icbm / vBarea) * 100;        
            }
            
        }       
    }
    else {
        if (!isNaN(vBcbm)) {
            if (vBcbm != 0) {
                vPer = (icbm / vBcbm) * 100;           
            }            
        }

    }

    if (!isNaN(vPer)) {

        document.getElementById(DSPper).innerHTML = vPer.toFixed(6);
        document.getElementById(HDPer).value = vPer.toFixed(6);
    }

    else {
        document.getElementById(DSPper).innerHTML = 0;
        document.getElementById(HDPer).value = 0;
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
        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.LOBH_WH.value);

        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + lQty_id + "|3|L, " + qty_id + "|3, " + lbatch_no + "|4|L, " + batch_no + "|4, " + lexp_date + "|5|L, " + exp_date + "|5, " + lmanu_date + "|6|L, " + manu_date + "|6 ");
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


function limitText(limitField, limitNum) {
    if (limitField.value.length > limitNum) {
        limitField.value = limitField.value.substring(0, limitNum);
    }
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
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td class="TITLE" colspan="4">
                <asp:label ID="lheader" runat ="server" /><input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" style="float:right" />
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
                        <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />   
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" Visible="false" />
                    <% End If%>
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="LOBH_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="LOBH_STATUS" runat="server" /></font>
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
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_BOOK_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:label ID="LOBH_BOOK_BY" runat="server" />
                </td>                
            </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_WH" runat="server" /></font>
                </td>
                <td id="Td1" colspan="3" runat="server">
                    <font size="2">
                    <asp:UpdatePanel runat="server" ID="WHUDP" RenderMode="Inline">
                        <ContentTemplate>
                             <asp:DropDownList ID="LOBH_WH" runat="server" AutoPostBack="true" CssClass="REQUIRED"></asp:DropDownList>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                        
                    </font>
                </td>
            </tr>       
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_START_DATE" runat="server" Text="Date:" /></font>
                </td>
                <td>
                    <asp:TextBox ID="LOBH_START_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="LOBH_START_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />  
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_RO_CODE" runat="server" Text="WPO No.:" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:Label runat="server" ID="LOBH_RO_CODE"  Width="100px" />
                        <asp:Image ID="Image_RO_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="center" />     
                    </font>
                </td>          
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="LBL_LOBH_END_DATE" runat="server" Text="End Date:" /></font>
                </td>
                <td>
                    <asp:TextBox ID="LOBH_END_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="LOBH_END_DATE" PopupButtonID="btnDATE_ID2" Format="dd/MM/yyyy" />  
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_PO_CODE" runat="server" Text="PO No.:" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:Label runat="server" ID="LOBH_PO_CODE" />
                    </font>
                </td>          
            </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_FD_DATE" runat="server" Text="Forecast Depart Date:" /></font>
                </td>
                <td>
                    <asp:TextBox ID="LOBH_FD_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="LOBH_FD_DATE" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />  
                </td>
                <td class="LabelTD" nowrap width="15%">
                     <font size="2">
                         <asp:Label ID="lbl_LOBH_AD_DATE" runat="server" Text="Actual Depart Date:" /></font>
                 </td>
                 <td>
                     <asp:TextBox ID="LOBH_AD_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                        <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                             ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                     <asp:CalendarExtender ID="CalendarExtender4" runat="server" CssClass="ajax_calendar" 
                             TargetControlID="LOBH_FD_DATE" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />  
                </td>
            </tr>
           
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_LOBH_REMARKS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="LOBH_REMARKS" runat="server" rows="5" Width="523px" 
                        TextMode="MultiLine"  onKeyDown="limitText(this,4000);" onKeyUp="limitText(this,4000);" onBlur="limitText(this,4000);" />
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
            <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1200px">
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
                    <asp:UpdatePanel runat="server" ID="btnDDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:Button ID="newrow" runat="server" Text="Add Item" CssClass="all_button" /> 
                        </ContentTemplate>
                    </asp:UpdatePanel>                    
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" visible ="false" />
                </td>
            </tr>            
            <tr>
                <td colspan="4">
                <asp:UpdatePanel runat="server" ID="GVUDP" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="LOBD_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                            <Columns>
                                 <asp:BoundField DataField="LOBD_SEQ" HeaderText="No." >
                                    <ControlStyle Width="30px"  />
                                    <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Stock No.">
                                    <ItemTemplate>
                                        <asp:textbox ID="LOBD_SKU_NO" runat="server" Font-Size="11px" MaxLength="80" width="98%" />
                                        <asp:hiddenfield runat="server" id="LOBD_ITM_CODE" />
                                        <asp:hiddenfield runat="server" id="mFlag" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="center" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pack Key" HeaderStyle-Width="80px">
                                    <ControlStyle Width="60px" />
                                    <ItemTemplate>
                                        <asp:textbox ID="PACK_KEY" runat="server" Font-Size="11px" Width="60px" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="center" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Item Name">
                                    <ItemTemplate> 
                                        <asp:textbox ID="LOBD_ITM_DESC" runat="server" Font-Size="11px" Width="98%" />
                                    </ItemTemplate>                                
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Loc.">
                                    <ItemTemplate>
                                           <asp:Label ID="dsp_LOBD_LOC" width="150px" runat="server" Font-Size="11px" CssClass="REQUIRED" />
                                           <asp:Image ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="right" />     
                                           <asp:HiddenField ID="LOBD_LOC" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="CBM">
                                <HeaderStyle Width="100px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="LOBD_CBM" runat="server" Font-Size="11px" Width="100px" style="text-align:right"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="%" HeaderStyle-HorizontalAlign="Right">
                                <HeaderStyle Width="100px" />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_LOBD_CBM_PERC" runat="server" Font-Size="11px" Width="80px" style="text-align:right"></asp:Label>
                                        <asp:HiddenField ID="LOBD_CBM_PERC" runat="server" />
                                        <asp:HiddenField ID="BN_UTILIZATION_TYPE" runat="server" />
                                        <asp:HiddenField ID="BN_CBM" runat="server" />
                                        <asp:HiddenField ID="BN_AREA" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Status">
                                    <ItemTemplate>
                                        <asp:Label ID="LOBD_STATUS" runat="server" Font-Size="11px" />                                    
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="center" />
                                </asp:TemplateField>
                                <asp:TemplateField ControlStyle-Width="50px">
                                    <ItemTemplate>
                                        <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                            CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                    </ItemTemplate>
                                    <ControlStyle Width="50px"/>
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
