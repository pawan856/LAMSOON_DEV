<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DRUMMain.aspx.vb" Inherits="OPERATION_STA_DRUMMain" %>
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

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_ILOC_BAL");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "wh", WH);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2, seqList");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}


function LocLookUp(seq, lb_id, hd_id, mWH) {
    //if (document.myform.editMode.value != "V") {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");

    if (seq == 1) {
        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.DRUM_WH_CODE.value);
//    } else {
//        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TR_WH_TO.value);

    }

    //setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedToWH('" + rowIdx + "')");
    setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

//    if (mWH == 'F') {
//        setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.TR_WH_FR.value);
//    }

//    if (mWH == 'T') {
//        setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.TR_WH_TO.value);
//    }

    setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
    document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
    document.hiddenForm.target = "locLookUp";
    document.hiddenForm.submit();
    //}
}


function maskKey(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
    return false;
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
        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.DRUM_WH_CODE.value);

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

function goToSL(storer_code, drum_id) {

    //alert("Alert");
    removeAllElementFromForm(document.hiddenForm);
    window.open("", "SLCable", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=384,left=80,top=80");

    setInterfaceDataToForm(document.hiddenForm, "DRUM_ID", drum_id);
    setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", storer_code);

    document.hiddenForm.action = "DRUM_SL.aspx";
    document.hiddenForm.target = "SLCable";
    document.hiddenForm.submit();
}


function goToPrint(storer_code, drum_id) {

    //alert("Alert");
    removeAllElementFromForm(document.hiddenForm);
    window.open("", "PrintMove", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=768,left=80,top=80");

    setInterfaceDataToForm(document.hiddenForm, "DRUM_ID", drum_id);
    setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", storer_code);

    document.hiddenForm.action = "MoveList.aspx";
    document.hiddenForm.target = "PrintMove";
    document.hiddenForm.submit();
}

function validNum(ctrl, num) {
    var hr;
    inval = ctrl.value
    if (inval != '') {
        if (inval > num) {
            ctrl.value = num;
        }
    }
}


function getLoad() {
    var load_modalPopup = $find('load_ModalPopupExtender');
    load_modalPopup.show();
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
    <asp:HiddenField ID="seqList" runat="server" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="4" class="TITLE" align="left">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:label ID="lheader" runat ="server" /></b>
                            </td>
                            <td width="100%" class="TITLE">
                                <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                    <input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                        <td class="menuTD" align="left">
                            <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                            <asp:Button ID="CancelBtn" Text ="Inactive" CssClass="all_button" runat ="server" />
                            <asp:Button ID="btnReOpen" Text="Re-Open" CssClass="all_button" runat="server" Visible="false" />
                            <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                                <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                                class="all_button" />
                        </td>
                        <td class="menuTD" align="right">
                            <asp:Button runat="server" ID="btnPrnLabel" text="Print Drum Label" CssClass="all_button" />
                            <asp:Button runat="server" ID="btnPrint" text="Print Movement List" CssClass="all_button" />
                        </td>
                    </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_ID" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:textbox ID="DRUM_ID" runat="server" MaxLength="80" CssClass="REQUIRED" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="DRUM_STATUS" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
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
                        <asp:Label ID="lbl_DRUM_DATE" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DRUM_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                       <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DRUM_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />  
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_WH_CODE" runat="server" /></font>
                </td>
                <td id="Td2" runat="server">
                    <font size="2">
                         <asp:DropDownList ID="DRUM_WH_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_TYPE" runat="server" /></font>
                </td>
                <td id="Td1" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="DRUM_TYPE" runat="server">
                        <asp:ListItem Text="DRUM" Value="DRUM" />
                        <asp:ListItem Text="TROLLEY" Value="TROLLEY" />
                        <asp:ListItem Text="OTHER" Value="OTHER" />
                    </asp:DropDownList>
                    </font>
                </td>         
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_LOC" runat="server" Text="Location" />:</font>
                </td>
                <td>
                     <asp:Label ID="dsp_DRUM_LOC" runat="server" />
                        <asp:HiddenField ID="DRUM_LOC" runat="server" />
                        <asp:HiddenField ID="emptyDrum_YN" runat="server" Value="N" />
                        <asp:Image ID="Image_Loc_LookUp_DRUM_LOC" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" Visible="false" />
                </td>
                 <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_IN_WH" runat="server" Text="In Warehouse?" /></font>
                </td>
                <td>
                    <asp:DropDownList runat="server" id="DRUM_IN_WH">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="Yes" Value="Y" />
                        <asp:ListItem Text="No" Value="N" />
                    </asp:DropDownList>
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <asp:Label runat="server" ID="lbl_DRUM_IS_RESERVED" Text="Is Reserved?" />    
                </td>
                <td>
                    <asp:CheckBox runat="server" ID="DRUM_IS_RESERVED" />
                </td>
                 <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_IS_EMPTY" runat="server" Text="is Empty Drum?" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="DRUM_IS_EMPTY" runat="server" Text="" /></font>
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <asp:Label runat="server" ID="Label1" Text="Under Inspection?" />    
                </td>
                <td colspan="3">
                    <asp:CheckBox runat="server" ID="DRUM_IS_INSP" />
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_SERIAL_NO" runat="server" Text="Serial No." />:</font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="DRUM_SERIAL_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>               
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_CLOSE_DATE" runat="server" Text="Close Date" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:label ID="DRUM_CLOSE_DATE" runat="server" />
                    </font>
                </td>               
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_TOTAL_PK_ITEM" runat="server" Text="Total Pickup Items" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:label ID="TOTAL_PK_ITEM" runat="server" />
                    </font>
                </td>                
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_REOPEN_BY" runat="server" text="Reopen By"/>:</font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:label ID="DRUM_REOPEN_BY" runat="server" />
                    </font>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_DESC" runat="server" Text="Description" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DRUM_DESC" runat="server" MaxLength="1000" Width="523px" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_DRUM_REMARKS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DRUM_REMARKS" runat="server" Height="117px" Width="523px" 
                        TextMode="MultiLine"></asp:TextBox>
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
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" Visible = "false" />

                    <asp:Button ID="btnShowSL" runat="server" Text="Show Short Length Cable" CssClass="all_button" />
                    <div style="float:right;display: inline; ">
                        <asp:Button ID="btnCollect" runat="server" Text="Supplier Collect" CssClass="all_button" />
                        <asp:Button ID="btnReturn" runat="server" Text="Return" CssClass="all_button" />
                        <asp:Button ID="btnRelocate" runat="server" Text="Relocate" CssClass="all_button" />
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="DRUD_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:TemplateField HeaderText="Date Time">
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="dsp_DRUD_DATE" Visible="false" />
                                    <asp:textbox ID="DRUD_DATE" runat="server" Font-Size="11px" width="60px" CssClass="REQUIRED"  />
                                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                                         ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                    <br />
                                    <asp:TextBox runat="server" ID="DRUD_DATE_HR" Font-Size="11px" onBlur="validNum(this,23);" width="30px" /> : 
                                    <asp:TextBox runat="server" ID="DRUD_DATE_MIN" Font-Size="11px" onBlur="validNum(this,59);" width="30px" />                                    
                                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                                         TargetControlID="DRUD_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                                    <asp:FilteredTextBoxExtender runat="server" ID="fil1" TargetControlID="DRUD_DATE_HR" FilterType="Numbers" />
                                    <asp:FilteredTextBoxExtender runat="server" ID="fil2" TargetControlID="DRUD_DATE_MIN" FilterType="Numbers" />
                                    <asp:HiddenField runat="server" ID="DRUD_MV_TYPE" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Movement">
                                <ItemTemplate>
                                    <asp:Label ID="DRUD_MOVEMENT" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Doc Type">
                                <ItemTemplate>
                                    <asp:dropdownlist ID="DRUD_DOC_TYPE" runat="server" Font-Size="11px" CssClass="REQUIRED">
                                       <asp:ListItem text="SELECT" Value="" />
                                       <asp:ListItem text="GR" Value="GR" />
                                       <asp:ListItem text="WIT" Value="WIT" />
                                       <asp:ListItem text="SI" Value="SI" />
                                       <asp:ListItem text="SR" Value="SR" />
                                       <asp:ListItem text="RFTI" Value="RFTI" />
                                       <asp:ListItem text="RFTR" Value="RFTR" />
                                       <asp:ListItem text="NA" Value="NA" />
                                    </asp:dropdownlist> 
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Doc No">
                                <ItemTemplate>
                                    <asp:textbox ID="DRUD_DOC_NO" runat="server" Font-Size="11px" Width="98%" MaxLength="80" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="By">
                                <ItemTemplate>
                                    <asp:textbox ID="DRUD_BY" runat="server" Font-Size="11px" MaxLength="80" Width="98%" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Warehouse">
                                <ItemTemplate>
                                    <asp:label ID="DRUD_WH" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Location">
                                <ItemTemplate>
                                   <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td class="REQUIRED">
                                                     <asp:Label ID="dsp_DRUD_LOC" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td class="REQUIRED">
                                                    <asp:Image ID="Image_DRUD_LOC" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="DRUD_LOC" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ItemTemplate>
                                    <asp:textbox ID="DRUD_REMARKS" runat="server" Font-Size="11px" MaxLength="4000" Width="98%" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
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
     <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
                ScrollBars="None">
                <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
                    height: 80px">
                    <tr>
                        <td align="center" style="background-color: White; width: 100%; height: 80px; vertical-align: middle">                            
                            <font color="#193B65" style="width: 100%; text-align: center; font-size: 16px;">Posting...please wait</font>
                        </td>
                    </tr>                    
                </table>
            </asp:Panel>
            <asp:HiddenField ID="loadDummy" runat="server" />
            <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
                TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
                DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
                Y="250">
            </asp:ModalPopupExtender>

    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
