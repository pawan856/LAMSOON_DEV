<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SRCloseMain.aspx.vb" Inherits="SRMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="ajx" %>
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

<div runat="server">
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

    function SRLookUp(STORER_CODE, SRC_WH) {
        if (STORER_CODE == '')
        {
            alert('Please select the Storer first!');
            return false;
        }

        if(document.myform.editMode.value!="V")
        {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "SRLookUP", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1024,height=768,left=5,top=15");
        
            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_SR");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "SRLookUP";
            document.hiddenForm.submit();
        }
    }

    function selectedItem()
    {   
            document.myform.moduleAction.value = "SELECTIM";
            document.myform.submit();
    }

    function WITLookUp(STORER_CODE) {
        if (STORER_CODE == '') {
            alert('Please select the Storer first!');
            return false;
        }

        if (SRC_WH == '') {
            alert('Please select the main warehouse first!');
            return false;
        }

        if (document.myform.editMode.value != "V") {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "WITLookUP", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1024,height=768,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_SRC_WIT");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedWIT()");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "witmList|1");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "WITLookUP";
            document.hiddenForm.submit();
        }
    }

    function selectedWIT() {
        document.myform.moduleAction.value = "SELECTWIT";
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

    function OpenItemLbls() {
           
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "srItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "SRC_code", document.getElementById("SRC_CODE_HF").value);
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
</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form id="myform" runat="server">
  <ajx:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
          <tr>
                <td class="TITLE" colspan="4">
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
                <td colspan="4" class="menuTD">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="menuTD" align="left">
                                <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                                <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_SRC'"
                                    class="all_button" />
                            </td>
                            <td class="menuTD" align="right">
                            <asp:button runat="server" id="btnPreweight" text="Pre-Weight" cssclass="all_button" />
                            <asp:button runat="server" id="btnUnPreweight" text="Un-Preweight" cssclass="all_button" />
                            <asp:UpdatePanel runat="server" id="bcloseUDP" RenderMode="Inline">
                                <ContentTemplate>
                                       <asp:button runat="server" id="btnClose" text="Close" cssclass="all_button" />
                                    </ContentTemplate>
                            </asp:UpdatePanel>
                                <asp:button runat="server" id="btnUnClose" text="Un-Close" cssclass="all_button" />
                                <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />
                                
                            </td>
                        </tr>
                    </table>                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap style="width:20%;">
                    <font size="2">
                        <asp:Label ID="lbl_SRC_CODE" runat="server" />:</font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="SRC_CODE" runat="server" />
                        <asp:HiddenField runat="server" ID="SRC_CODE_HF" />
                    </font>
                </td>
                <td class="LabelTD"  width="20%">
                    <font size="2">
                        <asp:Label ID="lbl_SRC_STATUS" runat="server" />:</font>
                </td>
                <td width="30%">
                    <font size="2">
                        <asp:label runat="server" ID="DSP_SRC_STATUS" />
                        <asp:hiddenfield ID="SRC_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" />:</font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" />
                    <asp:HiddenField runat="server" ID="STORER_CODE_HF" />
                    </font>
                </td>
               <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_CLOSE_DATE" runat="server"  Text="Closed Date" />:</font>
                </td>
                <td>
                    <asp:label ID="SRC_CLOSE_DATE" runat="server" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_DATE" runat="server" text="Date" />:</font>
                </td>
                <td>
                     <asp:TextBox ID="SRC_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                     <asp:ImageButton ID="btnDATE_ID2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <ajx:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="SRC_DATE" PopupButtonID="btnDATE_ID2" Format="dd/MM/yyyy" />  
                </td>                
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_CLOSE_BY" runat="server" Text="Closed By" />:</font>
                </td>
                <td id="Td1" runat="server">
                    <font size="2">
                        <asp:Label runat="server" ID="SRC_CLOSE_BY" />
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_WH" runat="server" Text="Warehouse" />:</font>
                </td>
                <td>
                     <asp:DropDownList runat="server" ID="SRC_WH" CssClass="REQUIRED" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_CANCEL_DATE" runat="server" Text="Cancelled Date" />:</font>
                </td>
                <td runat="server">
                    <asp:Label ID="SRC_CANCEL_DATE" runat="server" />
                </td>                
            </tr>           
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_CANCEL_BY" runat="server" Text="Cancelled By" />:</font>
                </td>
                <td id="Td2" runat="server" colspan="3">
                   <asp:Label runat="server" ID="SRC_CANCEL_BY" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_WIT_NO" runat="server" text="WIT No." />:</font>
                </td>
                <td colspan="3">
                     <asp:TextBox runat="server" ID="SRC_WIT_NO" MaxLength="80" Width="300px" />&nbsp;
                     <asp:UpdatePanel runat="server" id="WITUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:ImageButton runat="server" ID="btnWITLookup" ImageUrl="..\..\images\btn_search.gif" ImageAlign="Middle" />   
                        </ContentTemplate>
                     </asp:UpdatePanel>                     
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="Label1" runat="server" text="Total WIT Qty1" />:</font>
                </td>
                <td>
                     <asp:TextBox runat="server" ID="SRC_WIT_TOTAL_QTY" MaxLength="15" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="Label2" runat="server" Text="Total WIT KG" />:</font>
                </td>
                <td id="Td3" runat="server">
                   <asp:TextBox runat="server" ID="SRC_WIT_TOTAL_KG" MaxLength="15" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SRC_REM" runat="server" text="Remark" />:</font>
                </td>
                <td colspan="3">
                    <asp:TextBox runat="server" ID="SRC_REM" TextMode="MultiLine" Rows="3" width="500px"/>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td colspan="4" align="left">
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
                    <asp:updatepanel runat="server" ID="LOOKUPUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:Button ID="selectItemBtn" runat="server" Text="Select SR" CssClass="all_button" />
                        </ContentTemplate>
                    </asp:updatepanel> 
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <table style="width:100%; border-spacing:0; padding:0; background-color:transparent;">
                        <tr>
                            <td style="background-color:transparent;">No. of Stock Items Selected:&nbsp;<asp:label runat="server" ID="NO_OF_ITM_SKU" /></td>
                            <td style="background-color:transparent;">No. of SR Selected:&nbsp;<asp:label runat="server" ID="NO_OF_SR" /></td>
                            <td style="background-color:transparent; text-align:right;" width="400px">No. of SR Selected  which did not Complete Approval:</td>
                            <td width="100px" runat="server" id="ISR_COUNT_TD">&nbsp;<asp:label runat="server" ID="NO_OF_ISR" /></td>
                        </tr>
                        <tr>
                            <td style="background-color:transparent;">Total Selected Qty1:&nbsp;<asp:label runat="server" ID="TOTAL_QTY1" /></td>
                            <td style="background-color:transparent;" colspan="3">Total Selected KG:&nbsp;<asp:label runat="server" ID="TOTAL_KG" /></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="seq_no" 
                        HorizontalAlign="Left">
                        <Columns>
                            <asp:TemplateField HeaderText="No." HeaderStyle-Width="30px">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:TextBox ID="SEQ_NO" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                    <asp:hiddenfield runat="server" id="RT_CODE" />
                                </ItemTemplate>
                            </asp:TemplateField>                          
                            <asp:BoundField HeaderText="Stock Return No." ItemStyle-Font-Size="11px" DataField="RT_CODE" />
                            <asp:BoundField HeaderText="Warehouse" ItemStyle-Font-Size="11px" DataField="RT_WH" />
                            <asp:BoundField HeaderText="Credit Form No" ItemStyle-Font-Size="11px" DataField="RT_REF_NO2" />
                            <asp:BoundField HeaderText="Creadit Form Receive Date" ItemStyle-Font-Size="11px" DataField="RT_DATE" />
                            <asp:TemplateField HeaderText="Status" ItemStyle-Font-Size="11px">
                                <ItemTemplate>
                                <asp:Label runat="server" ID="RT_STATUS" />
                                </ItemTemplate>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                            <asp:BoundField HeaderText="Stock No." ItemStyle-Font-Size="11px" DataField="itm_sku_no_str" />
                            <asp:BoundField HeaderText="Location" ItemStyle-Font-Size="11px" DataField="location_str" />
                            <asp:BoundField HeaderText="Total Qty1" ItemStyle-Font-Size="11px" DataField="RTD_RCV_QTY" ItemStyle-HorizontalAlign="Right"  DataFormatString={0:N0} />
                            <asp:BoundField HeaderText="Total KG" ItemStyle-Font-Size="11px" DataField="RTD_KG" ItemStyle-HorizontalAlign="Right" />
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Remove" CssClass="all_button" Font-Bold="false" />
                                </ItemTemplate>
                                <HeaderStyle Width="70px" />
                            </asp:TemplateField>
                        </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" value="Back"  onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_SRC'"
                        class="all_button" />
                </td>
            </tr>
        </table>

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
         <ajx:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
             TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
             DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
             Y="250">
         </ajx:ModalPopupExtender>

        <asp:UpdatePanel runat="server" ID="UDPClose">
            <ContentTemplate>
                <asp:Panel ID="pnlClose" runat="server" CssClass="modalPopup" style="display: none" >
                     <table width="800px">
                         <tr class="TITLE">
                             <td colspan="2" class="TITLE">
                                 <asp:Panel runat="Server" ID="PanelDrag" Style="cursor: move;">
                                   <font size="2">Confirmation:</font> </asp:Panel></td>
                         </tr>
                         <tr runat="server" id="TR_CAUTION">
                            <td>
                                Please note the following(s):
                                <asp:BulletedList runat="server" ID="cautionList" BulletStyle="Disc" BackColor="White" />                                
                            </td>
                         </tr>
                         <tr>
                            <td>
                              <b>Confirm to Close out?</b>
                            </td>
                         </tr>
                         <tr>
                             <td align="left" colspan="2">
                                 <asp:Button runat="server" ID="btnPOK" CssClass="all_button" text="OK" />&nbsp; <asp:Button runat="server" ID="btnPCancel"  CssClass="all_button" Text="Cancel" />
                             </td>
                         </tr>
                     </table> 
        </asp:Panel>            
        <ajx:ModalPopupExtender ID="btnClose_ModalPopupExtender" runat="server" 
                    DynamicServicePath="" 
                    Enabled="True" 
                    TargetControlID="dummy" 
                    PopupControlID="pnlClose"
                    BackgroundCssClass="modalBackground"
                    DropShadow="true"         
                    CancelControlID="btnPCancel"
                    RepositionMode="RepositionOnWindowResize"
                    PopupDragHandleControlID="PanelDrag"
                    y="100" >                    
        </ajx:ModalPopupExtender>     
        <asp:HiddenField ID="dummy" runat="server" />    
            </ContentTemplate>
            <Triggers>
                <asp:PostBackTrigger runat="server" ControlID="btnPOK" />
            </Triggers>
        </asp:UpdatePanel>
    </div>
    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="witmList" runat ="server" />
    <asp:HiddenField ID="new_action" runat ="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
