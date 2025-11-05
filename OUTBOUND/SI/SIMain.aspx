<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SIMain.aspx.vb" Inherits="OUTBOUND_SI_SIMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Stock Issue</title>
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

function LocLookUp(lb_id, hd_id) {
    if (document.myform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");

        setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.IS_WH.value);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function ItemLookUp(STORER_CODE,WH) {
    if (STORER_CODE == '')
    {
    
    }

    if (WH == '') 
    {
        alert('Please select the Warehouse first!');
        return false;
    }

    if(document.myform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,resizable=yes,width=1200,height=800,left=5,top=15");

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

function selectedItem(){   
        document.myform.moduleAction.value = "SELECTIM";
        document.myform.submit();
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


    function removeClass(val) {
        if (val == 'WRITEOFF') {
            document.getElementById('IS_PLAN_DATE').className = "";
        }
        else {
            document.getElementById('IS_PLAN_DATE').className = "REQUIRED";
        }
    }

    function getLoad() {
        var load_modalPopup = $find('load_ModalPopupExtender');
        load_modalPopup.show();
    }

$(document).keypress(
    function(event){
     if (event.which == '13') {
        event.preventDefault();
      }


});

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:label ID="lheader" runat ="server" />
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0" Visible="false"></asp:Table>
                    <div style="float:right; display:inline;"><input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" /></div>
                </td>
            </tr>
            
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=OB_SI'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                    <%If Session("pagemode") <> "N" Then%>             
                        <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />
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
                        <asp:Label ID="lbl_IS_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="IS_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="IS_STATUS" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack = "true" ></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_TYPE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="IS_TYPE" runat="server" onchange="javascript:removeClass(this.value)"></asp:DropDownList>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="IS_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="IS_DATE" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_ISSUED_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="IS_ISSUED_BY" runat="server" MaxLength="20"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_PLAN_DATE" runat="server" /></font>
                </td>
                <td width="150px" colspan="3">
                    <asp:TextBox ID="IS_PLAN_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" ClientIDMode="Static"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="IS_PLAN_DATE" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_RETURN_DATE" runat="server" /></font>
                </td>
                <td width="150px" colspan="1">
                    <asp:TextBox ID="IS_RETURN_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="IS_RETURN_DATE" PopupButtonID="ImageButton3" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_RETURN_DOC_NO" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_RETURN_DOC_NO" runat="server" 
                        style="text-align:right" onkeypress="return maskKey(event);" Width="92px"></asp:TextBox>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_PROJECT_NO" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="IS_PROJECT_NO" runat="server" MaxLength="20"></asp:DropDownList>
                    <asp:Label ID="lbl_PRJ_NAME" runat="server" />
                    </font>
                </td>
            <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_TOT_PALLET" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_TOT_PALLET" runat="server" MaxLength="12" 
                        style="text-align:right" onkeypress="return maskKey(event);" Width="92px"></asp:TextBox>
                    </font>
                </td>
            </tr>
            <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_REQ_BY" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_REQ_BY" runat="server" MaxLength="30"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_REF_NO" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_REF_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_REQ_TEL" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_REQ_TEL" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_REQ_EMAIL" runat="server" /></font>
                </td>
                <td nowrap >
                    <font size="2">
                        <asp:TextBox ID="IS_REQ_EMAIL" runat="server" MaxLength="30" Width="240px"></asp:TextBox>
                    </font>
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_WH" runat="server" /></font>
                </td>
                <td width="26%" runat="server">
                    <font size="2">
                         <asp:DropDownList ID="IS_WH" runat="server" ></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_BATCH_NO" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                      <asp:TextBox ID="IS_BATCH_NO" MaxLength="20" runat="server" /></font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_CABLE_YN" runat="server" text="Is Cable:" /></font>
                </td>
                <td colspan="3">
                    <asp:CheckBox runat="server" ID="IS_CABLE_YN" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_IS_REM" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="IS_REM" runat="server" Height="117px" Width="523px" 
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
                <td colspan="4" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:UpdatePanel runat="server" id="LOOKUPUDP" RenderMode="Inline" >
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
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="ISD_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                             <asp:BoundField DataField="ISD_SEQ" HeaderText="No." >
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Item Code">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="isd_itm_code" runat="server" Font-Size="11px"></asp:label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="isd_pack_key" runat="server" Font-Size="11px"></asp:label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="isd_itm_name" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Stock No.">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="itm_SKU_NO" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Batch No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="isd_batch_no" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="isd_pallet_no" runat="server" Width="60px" Font-Size="11px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Carton No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="isd_carton_no" runat="server" Width="60px" Font-Size="11px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Ref. No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="isd_ref_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="isd_issue_qty" runat="server" Font-Size="11px" Width="30px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Loc">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="isd_loc" runat="server"></asp:DropDownList>
                                       <%-- <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_isd_loc" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td align="right">
                                                    <asp:Image ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="isd_loc" runat="server" />--%>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Cut Y/N">
                                <HeaderStyle HorizontalAlign="Left" Width="40px" />
                                <ItemTemplate>
                                    <asp:CheckBox runat="server" ID="ISD_CUT_YN" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Short Length">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label runat="server" ID="ISD_SL" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label runat="server" ID="ISD_UOM2" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty 2">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="ISD_QTY2" runat="server" Font-Size="11px" Width="30px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Serial No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="isd_serial_no" runat="server" Font-Size="11px" Width="50px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="ISD_rem" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
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
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=OB_SI'"
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
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />
    <asp:HiddenField ID="seqList" runat ="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
