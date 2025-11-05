<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SRMain.aspx.vb" Inherits="SRMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <title>Stock Return</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script src="../../js/jquery-1.9.1.min.js"></script>
<script src="../../js/jquery-migrate-1.1.1.min.js"></script> 

<script language="javascript">
//function DisableDeleteButton() {
//    var grp = document.getElementsByName("btnDelete");
//    var count;
//    count = grp.length;
//    if (count == 1) {
//       grp[0].disabled = true;
//    }
//}

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

function checkKey(objEvent) {
    alert(objEvent.keyCode);
}

function VendorLookUp() {

    if(document.myform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "vendLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_VEND");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "dsp_VND_CODE||L, VND_CODE, dsp_VND_NAME|1|L, VND_NAME|1");
        setInterfaceDataToForm(document.hiddenForm, "sc", document.myform.STORER_CODE.value);
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "vendLookUp";
        document.hiddenForm.submit();
    }
}

function LocLookUp(lb_id, hd_id) {
    if(document.myform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.RT_WH.value);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function DrumLookUp(hd_id, hd_id2) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "drumLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_DRUM");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", hd_id + ", " + hd_id2 + "|1");
        setInterfaceDataToForm(document.hiddenForm, "sc", document.myform.STORER_CODE.value);
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "drumLookUp";
        document.hiddenForm.submit();
    }
}

function ItemLookUp(STORER_CODE) {
    if (STORER_CODE == '')
    {
        alert('Please select the Storer first!');
        return false;
    }

    if(document.myform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=700,height=500,left=5,top=15");
        
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

function selectedItem()
{   
        document.myform.moduleAction.value = "SELECTIM";
        document.myform.submit();
}

    function checkSB(STORER_CODE) {
        //if(document.myform.editMode.value!="V")
        //{
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "sbLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=1200,height=500,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "menu_code", "INQ_001");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "ctemp", true);
            setInterfaceDataToForm(document.hiddenForm, "screadonly", true); 

            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "sbLookUp";
            document.hiddenForm.submit();
        //}
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

    function OpenItemLbls() {
           
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "srItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "rt_code", document.getElementById("RT_CODE_HF").value);
        setInterfaceDataToForm(document.hiddenForm, "so_date", document.getElementById("SO_DATE_HF").value);
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE_HF").value);
            
        document.hiddenForm.action = "SR_LABEL/item_label_print.aspx";
        document.hiddenForm.target = "srItemLbls";
        document.hiddenForm.submit();
    }

    function getLoad() {
        var load_modalPopup = $find('load_ModalPopupExtender');
        load_modalPopup.show();
    }

    function NewBtn() {
        if (document.getElementById("editMode").value != "V") {
            if (confirm('Do you want to save your data before directing to a new record?'))
                document.myform.new_action.value = "Y";
            else
                document.myform.new_action.value = "N";            
        }
        else {
            document.myform.new_action.value = "N";
        }
        __doPostBack('btnNew', '');
    }

$(document).keypress(
    function(event){
     if (event.which == '13') {
        event.preventDefault();
      }


});



</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form id="myform" runat="server">
  <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td class="TITLE" colspan="8">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lheader" runat="server" /></b>
                            </td>
                            <td width="100%" class="TITLE">
                                <input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="8" class="menuTD">
                F<table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button runat="server" id="btnNew" Text="New" CssClass="all_button" OnClientClick="NewBtn();" /></br>
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <asp:Button ID="saveBtn4" runat="server" Text="Save" CssClass="all_button" Visible = "false" />                        
                        <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_SR'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <div style="display:inline;float:left;"><asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" /></div>
                        <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />
                        <asp:button runat="server" ID="btnSubmit" Text="Submit for Approval" CssClass="all_button" />
                        <asp:button runat="server" ID="btnUnSubmit" Text="Un-Submit for Approval" CssClass="all_button" Visible = false  OnClientClick="return confirm(&quot;Are you sure to UnSubmit this Record?&quot;);" />
                        <asp:UpdatePanel runat="server" ID="UDPBTNAPP" RenderMode="Inline" >
                            <ContentTemplate>
                                <asp:button runat="server" ID="btnApprove" Text="Approve" CssClass="all_button" />
                            </ContentTemplate>    
                        </asp:UpdatePanel>
                        <input type="button" runat="server" ID="btnPringLbl" value="Print Item Label" Class="all_button" onclick="Javascript:OpenItemLbls();" />
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server"/>
                        <asp:Button ID="btnUnPost" Text="Un-Post" CssClass="all_button" runat="server" />
                        
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="RT_CODE" runat="server" />
                        <asp:HiddenField runat="server" ID="RT_CODE_HF" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:label runat="server" ID="DSP_RT_STATUS" />
                        <asp:hiddenfield ID="RT_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" />
                    <asp:HiddenField runat="server" ID="STORER_CODE_HF" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_TYPE" runat="server" /></font>
                </td>
                <td colspan="3" runat="server">
                    <font size="2">
                    <asp:DropDownList ID="RT_TYPE" runat="server" />
                    </font>
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="RT_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RT_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" /> 
                     <asp:HiddenField runat="server" ID="SO_DATE_HF" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_RCV_BY" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RT_RCV_BY" runat="server" MaxLength="20"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_CONS_CODE" runat="server" /></font>
                </td>
                <td width="150px">
                     <asp:TextBox ID="RT_CONS_CODE" runat="server" MaxLength="20"></asp:TextBox>
                   
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_CUS_CODE" runat="server" /></font>
                </td>
                <td colspan="3" runat="server">
                    <asp:TextBox ID="RT_CUS_CODE" runat="server" MaxLength="20"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_TOT_PALLET" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="RT_TOT_PALLET" runat="server" MaxLength="12" style="text-align:right"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_C8_YN" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:Label ID="RT_C8_YN" runat="server" /></font> &nbsp; <asp:Button ID="BTN_C8_YN" Text ="COMPLETE" CssClass="all_button" runat ="server" />&nbsp;<asp:Button ID="BTN_UNC8_YN" Text ="UNCOMPLETE" CssClass="all_button" runat ="server" />
                </td>      
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="RT_BY" runat="server" MaxLength="30"></asp:TextBox>
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_REF_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RT_REF_NO" runat="server" MaxLength="30"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_BY_TEL" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="RT_BY_TEL" runat="server" MaxLength="80"></asp:TextBox>
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_BY_EMAIL" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RT_BY_EMAIL" runat="server" MaxLength="30"></asp:TextBox>
                </td>  
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_WH" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="RT_WH" runat="server" MaxLength="20" cssclass="REQUIRED" AutoPostBack="True" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_BATCH_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RT_BATCH_NO" runat="server" MaxLength="20"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_REF_NO2" runat="server" Text="Credit Forms No." />:</font>
                </td>
                <td>
                    <font size="2">
                    <asp:UpdatePanel runat="server" RenderMode="Inline" id="cdUDP">
                        <ContentTemplate>
                        <table width="100%">
                            <tr>
                                <td runat="server" id="td1">
                                     <asp:TextBox runat="server" ID="RT_REF_NO2" MaxLength="120"/>
                                   <%-- <asp:TextBox runat="server" ID="RT_REF_NO2" MaxLength="120" CssClass="REQUIRED" AutoPostBack="true" />--%>        
                                </td>
                            </tr>
                        </table>
                        </ContentTemplate>
                    </asp:UpdatePanel>                        
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_REF_DOC_NO" runat="server" text="RMA Num" />:</font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RT_REF_DOC_NO" runat="server" MaxLength="20" />
                    <asp:Button runat="server" ID="btnGetSI" Text="Get SI Items" CssClass="all_button" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_APPROVE_CODE" runat="server" Text="Approval Code" />:</font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:label runat="server" ID="RT_APPROVE_CODE" />
                    </font>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="RT_REM" runat="server" Height="117px" Width="484px" TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RT_FAULT_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="RT_FAULT_REM" runat="server" Height="117px" Width="484px" TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="8">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td colspan="8" align="left" width="100%">
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
                <td class="TITLE" colspan="6">
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
                <td colspan="8" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:Button ID="btnCopyItem" runat="server" CssClass="all_button" Text="Copy Items" />
                    <asp:updatepanel runat="server" ID="LOOKUPUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" />
                            &nbsp;
                        </ContentTemplate>
                    </asp:updatepanel> 
                </td>
            </tr>
            <tr>
                <td colspan="8">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="rtd_seq" 
                        HorizontalAlign="Left">
                        <Columns>
                            <asp:TemplateField ControlStyle-Width="30px">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cb_copy" runat="server" />
                                </ItemTemplate>
                                <ControlStyle Width="30px"></ControlStyle>
                                <HeaderStyle Width="30px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="No.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_seq" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Code">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="ITM_SKU_NO" Font-Size="11px" />
                                    <asp:hiddenfield ID="rtd_itm_code" runat="server" />
                                    <asp:hiddenfield ID="ITM_SERIAL_NO_YN" runat="server" />
                                    <asp:hiddenfield ID="RTD_LOC_WH" runat="server" />
                                    <asp:hiddenfield ID="SAP_MAT_DOC_NO" runat="server" />
                                    <asp:hiddenfield ID="SAP_MAT_DOC_ITEM" runat="server" />
                                    <asp:hiddenfield ID="RTD_STATUS" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="rtd_pack_key" runat="server" Font-Size="11px" ></asp:label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Lot">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                       <asp:TextBox ID="rtd_batch_no" runat="server" Width="70px" Font-Size="11px" MaxLength="20"></asp:TextBox>
                                   <%--  <asp:ComboBox ID="rtd_batch_no" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>--%>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_pallet_no" runat="server" Width="70px" Font-Size="11px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Carton No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_carton_no" runat="server" Width="70px" Font-Size="11px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Vendor Code">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dsp_rtd_vnd_code" runat="server" Font-Size="11px" />
                                    <asp:HiddenField ID="rtd_vnd_code" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="rtd_itm_name" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Reference">
                                <ControlStyle Width="90px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_ref_no" runat="server" Font-Size="11px" Width="90px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Expiry Date">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="RTD_EXPIRY_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                     <asp:ImageButton ID="btnCal01" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calEXP" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="RTD_EXPIRY_DATE" PopupButtonID="btnCal01" Format="dd/MM/yyyy" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Manufactory Date">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="RTD_MANU_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                     <asp:ImageButton ID="btnCal02" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calMANU" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="RTD_MANU_DATE" PopupButtonID="btnCal02" Format="dd/MM/yyyy" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            
                            <asp:TemplateField HeaderText="Ori. Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="RTD_KG" ReadOnly="true" runat="server" Font-Size="11px" Width="40px" style="text-align:right"  onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Ori. UOM">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="RTD_DRUM_LV" ReadOnly="true" runat="server" Font-Size="11px" Width="100px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Subinventory">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemTemplate>
                                    <asp:TextBox ID="RTD_WH" runat="server" Font-Size="11px" Width="40px" AutoPostBack="True" OnTextChanged="RTD_WH_TextChanged"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" Width="200px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Loc">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemTemplate>
                                     <%--<asp:DropDownList ID="rtd_loc" runat="server"></asp:DropDownList>--%>
                                    <asp:ComboBox ID="rtd_loc" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                                      <%--  <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td class="REQUIRED">
                                                     <asp:Label ID="dsp_rtd_loc" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td>
                                                   <div style="float:right;display:inline;"><asp:Image ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" /></div>
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="rtd_loc" runat="server" />--%>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" Width="200px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_rcv_qty" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="ITM_UOM" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:TextBox>
                                    <%--<asp:Label runat="server" ID="ITM_UOM" Font-Size="11px" />--%>
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Serial No." Visible="false">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rtd_serial_no" runat="server" Font-Size="11px" Width="100px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Drum ID"  Visible="false">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>                                    
                                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td>
                                                     <asp:TextBox ID="RTD_DRUM_ID" runat="server" Font-Size="11px" Width="100px" MaxLength="50"></asp:TextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            
                            <asp:TemplateField HeaderText="Qty2">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="RTD_QTY2" runat="server" Font-Size="11px" Width="40px" style="text-align:right"  onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="RTD_UOM2" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:TextBox>
                                    <%--<asp:Label runat="server" ID="RTD_UOM2" Font-Size="11px" />--%>
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
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="8" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <asp:Button ID="saveBtn3" runat="server" Text="Save" CssClass="all_button" Visible = "false" />
                    <input id="btnBack" type="button" value="Back"  onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_SR'"
                        class="all_button" />
                </td>
            </tr>
        </table>

        <asp:UpdatePanel runat="server" ID="UDPApp">
            <ContentTemplate>
                <asp:Panel ID="pnlAPPRV" runat="server" CssClass="modalPopup" style="display: none">
                     <table width="800px">
                         <tr class="TITLE">
                             <td colspan="2" class="TITLE">
                                 <asp:Panel runat="Server" ID="PanelDrag" Style="cursor: move;">
                                   <font size="2">Please Enter the Approval Code to approve the stock return:</font> </asp:Panel></td>
                         </tr>
                         <tr>
                             <td class="LabelTD" width="100px">
                                <asp:Label runat="server" ID="lbl_APPROV" Text="Approval Code:" />
                             </td>
                             <td>
                                <asp:textbox runat="server" ID="pnl_RT_APPROVE_CODE" MaxLength="80" />
                             </td>
                         </tr>
                         <tr>
                             <td align="left" colspan="2">
                                 <asp:Button runat="server" ID="btnPOK" CssClass="all_button" text="OK" />&nbsp; <asp:Button runat="server" ID="btnPCancel"  CssClass="all_button" Text="Cancel" />
                             </td>
                         </tr>
                     </table> 
        </asp:Panel>            
        <asp:ModalPopupExtender ID="btnApprove_ModalPopupExtender" runat="server" 
                    DynamicServicePath="" 
                    Enabled="True" 
                    TargetControlID="dummy" 
                    PopupControlID="pnlAPPRV"
                    BackgroundCssClass="modalBackground"
                    DropShadow="true"         
                    CancelControlID="btnPCancel"
                    RepositionMode="RepositionOnWindowResize"
                    PopupDragHandleControlID="PanelDrag" >
        </asp:ModalPopupExtender>     
        <asp:HiddenField ID="dummy" runat="server" />    
            </ContentTemplate>
            <Triggers>
                <asp:PostBackTrigger runat="server" ControlID="btnPOK" />
            </Triggers>
        </asp:UpdatePanel>

    <!-- Update panel that decided for validation alert -->

        <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>

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
    </div>

    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />
    <asp:HiddenField ID="qtyList" runat ="server" />
    <asp:HiddenField ID="new_action" runat ="server" />
    </form>

    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>

</html>
