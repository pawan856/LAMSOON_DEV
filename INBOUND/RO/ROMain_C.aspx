<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ROMain_C.aspx.vb" Inherits="INBOUND_RO_ROMain_C" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
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
        if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 45) || (iKeyCode == 40) || (iKeyCode == 41)) return true;
        return false;
    }

    function maskNumOnly(objEvent) {
        var iKeyCode;
        iKeyCode = objEvent.keyCode;
        if (iKeyCode >= 48 && iKeyCode <= 57) return true;
        return false;
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

            window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");
            
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
        if(document.myform.editMode.value!="V")
        {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "sbLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=1200,height=500,left=5,top=15");
            
            setInterfaceDataToForm(document.hiddenForm, "menu_code", "INQ_001");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "noreset", "Y");
            setInterfaceDataToForm(document.hiddenForm, "ctemp", true);
            setInterfaceDataToForm(document.hiddenForm, "screadonly", true); 


            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "sbLookUp";
            document.hiddenForm.submit();
        }
    }
    
    function sumPCStotal() {
        var DataGridObj = document.getElementById('GridView1');

        var tot = 0;
        var eachVal;

        if (DataGridObj.rows.length > 0) {
            for (var iRow = 1; iRow <= DataGridObj.rows.length - 1; iRow++) {
                if (iRow == DataGridObj.rows.length -1) {
                    DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].innerText = tot;
                    DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].align = "Right";
                    DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].style.fontWeight = "bold";
                } else {
                    eachVal = 0

                    if (DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].children[0].value != "") {
                        eachVal = parseFloat(DataGridObj.rows[iRow].cells[<%=ViewState("cIndex")%>].children[0].value.replace(/,/gi, ""));
                    }
                        tot = parseFloat(tot) + parseFloat(eachVal);
                    }
            }
        } 
    }

    function OpenItemLbls() {
        
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "roItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "ro_code", document.getElementById("RO_CODE_HF").value);
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState("STORER_CODE")%>");
    }
    
    document.hiddenForm.action = "RO_LABELS/item_label_print.aspx";
    document.hiddenForm.target = "roItemLbls";
    document.hiddenForm.submit();
}


    function OpenOrd() {
        
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "roPrint", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=0,width=1200,height=800,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "ro_code", document.getElementById("RO_CODE_HF").value);
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState("STORER_CODE")%>");
    }
    
    document.hiddenForm.action = "RO_RECEIPT/stock_receipt.aspx";
    document.hiddenForm.target = "roPrint";
    document.hiddenForm.submit();
}

    </script>
    </div>    
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" >
    <br />
    <form id="myform" runat="server">
     <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />
    <asp:HiddenField ID="qtyList" runat ="server" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
            <tr>
                <td colspan="6" class="TITLE">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td class="TITLE" align="left">
                            Replenishmen Order
                        </td>
                        <td class="TITLE" align="right">
                           
                        </td>
                    </tr>
                  </table>
                </td>
            </tr>
            <asp:HiddenField ID="lheader" runat ="server" />
            <tr>
                <td colspan="6" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=IB_RO_C <%else %>../../IMPORT/UL_RO/UploadRO.aspx<%end if %>'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" />                        
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:Label ID="RO_CODE" runat="server" />
                        <asp:hiddenField ID="RO_CODE_HF" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="3">
                        <asp:Label ID="RO_STATUS" runat="server" /></font>
                        <asp:hiddenfield ID="STORER_CODE" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_DATE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="RO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                     <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />    
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_TYPE" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:TextBox ID="RO_TYPE" runat="server" MaxLength="20" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_REF_NO" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:TextBox ID="RO_REF_NO" runat="server" MaxLength="30" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_BATCH_NO" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="RO_BATCH_NO" runat="server" Width ="130px" MaxLength="20"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_ETD" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="RO_ETD" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_ETD" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_ETA" runat="server" /></font>
                </td>
                <td colspan="6">
                    <asp:TextBox ID="RO_ETA" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_ETA" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />
                </td>
                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_RO_CTRY_ORIGIN" runat="server" Text="Origin" /></font>:
                </td>
                <td nowrap colspan="5">
                    <font size="3">
                        <asp:DropDownList runat="server" ID="RO_CTRY_ORIGIN" Font-Size="12px"/>
                    </font>
                </td>
                
            </tr>   
            <tr>
                 <td class="LabelTD">
                    <font size="3">
                        <asp:Label ID="lbl_RO_SHIP_MODE" runat="server" text="Ship Mode" />:
                    </font>
                </td>
                <td colspan="5">
                    <asp:DropDownList runat="server" ID="RO_SHIP_MODE" />
                </td>
            </tr>       
            <tr>
                <td class="LabelTD" nowrap valign="top">
                    <font size="3">
                        <asp:Label ID="lbl_RO_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="RO_REM" runat="server" Height="117px" Width="422px" 
                        TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td colspan="6" align="left" width="100%">
                    <table border="1" cellspacing="0" cellpadding="0" width="100%">
                        <tr>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="3">
                                    <asp:Label ID="lbl_sys_cb" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="3">
                                    <asp:Label ID="sys_cb" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="3">
                                    <asp:Label ID="lbl_sys_lub" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="3">
                                    <asp:Label ID="sys_lub" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="3">
                                    <asp:Label ID="lbl_sys_cd" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="3">
                                    <asp:Label ID="sys_cd" runat="server" />
                                </font>
                            </td>
                            <td class="LabelTD" width="10%" align="right" nowrap>
                                <font size="3">
                                    <asp:Label ID="lbl_sys_lud" runat="server" />: </font>
                            </td>
                            <td class="sysInfoBorder" width="15%">
                                <font size="3">
                                    <asp:Label ID="sys_lud" runat="server" />
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Panel runat="server" ID="unusedFileds" Visible="false">
           <asp:Label ID="lbl_RO_RCV_BY" runat="server" Visible="false" />
           <asp:TextBox ID="RO_RCV_BY" runat="server" MaxLength="20" Width="160px" Visible="false" />
           <asp:Label ID="lbl_RO_ISSUED_BY" runat="server" />
           <asp:TextBox ID="RO_ISSUED_BY" runat="server" Width ="160px" MaxLength="50"></asp:TextBox>
           <asp:Label ID="lbl_RO_CUST_INV_NO" runat="server" text="Customer Invoice No." />:
           <asp:TextBox runat="server" id="RO_CUST_INV_NO" MaxLength="50" />
           <asp:Label ID="lbl_RO_STORER_TEL" runat="server" />
           <asp:TextBox ID="RO_STORER_TEL" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
           <asp:Label ID="lbl_RO_STORER_FAX" runat="server" /></font>
           <asp:TextBox ID="RO_STORER_FAX" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
           <asp:Label ID="lbl_RO_STORER_EMAIL" runat="server" />
           <asp:TextBox ID="RO_STORER_EMAIL" runat="server" MaxLength="30"></asp:TextBox>
           <asp:Label ID="lbl_RO_STORER_CONT" runat="server" />
           <asp:TextBox ID="RO_STORER_CONT" runat="server" MaxLength="30"></asp:TextBox>
           <asp:Label ID="lbl_RO_DEST" runat="server" />
           <asp:TextBox ID="RO_DEST" runat="server" MaxLength="50"></asp:TextBox>
           <asp:Label ID="lbl_RO_SEAL_NO" runat="server" text="Seal No." />
           <asp:TextBox runat="server" id="RO_SEAL_NO" MaxLength="100" />
           <asp:Label ID="lbl_RO_CONTAINER_NO" runat="server" text="Container No." />
           <asp:TextBox runat="server" id="RO_CONTAINER_NO" MaxLength="50" />
           <asp:Label ID="lbl_RO_TRACK_NO" runat="server" /></font>
           <asp:TextBox ID="RO_TRACK_NO" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
           <asp:Label ID="lbl_STORER_CODE" runat="server" Visible="false" />
           <asp:Label ID="lbl_PRJ_CODE" runat="server" />
           <asp:DropDownList ID="PRJ_CODE" runat="server" MaxLength="20"></asp:DropDownList>
        </asp:Panel>
        <table border="0" cellspacing="1" cellpadding="1" style="width:100%">
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
                <td colspan="6" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" UseSubmitBehavior ="false" />
                    <asp:Button ID="btnCopyItem" runat="server" Text="Copy Items" CssClass="all_button" UseSubmitBehavior ="false" />
                </td>
            </tr>
            <tr>
                <td colspan="6">
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="ro_code" HorizontalAlign="Left" ShowFooter="true">
                           <RowStyle CssClass="GV" />
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
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_disp_seq" runat="server" Font-Size="11px" style="text-align:center" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Vendor Code">
                                <ControlStyle Width="120px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dsp_rod_vnd_code" runat="server" Font-Size="11px" Width="120px"></asp:Label>
                                    <asp:HiddenField ID="rod_vnd_code" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Code">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="rod_itm_code" runat="server" Font-Size="11px" Width="150px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Stock No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" Width="150px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="150px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="rod_itm_name" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Description">
                                <ControlStyle Width="150px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_desc" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Pack Key">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="rod_pack_key" runat="server" Width="80px" Font-Size="10px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_pallet_no" runat="server" Width="70px" Font-Size="11px"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Carton No">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_carton_no" runat="server" Font-Size="11px" Width="70px"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:ComboBox ID="rod_batch_no" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                                    <asp:hiddenfield runat="server" ID="hd_rod_batch_no" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Reference">
                                <ControlStyle Width="80px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_ref_no" runat="server" Font-Size="11px" Width="80px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Expiry Date">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="ROD_EXPIRY_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                     <asp:ImageButton ID="btnCal01" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calEXP" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="ROD_EXPIRY_DATE" PopupButtonID="btnCal01" Format="dd/MM/yyyy" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Manufactory Date">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="ROD_MANU_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                     <asp:ImageButton ID="btnCal02" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calMANU" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="ROD_MANU_DATE" PopupButtonID="btnCal02" Format="dd/MM/yyyy" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Parent">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_itm_parent" runat="server" Font-Size="11px" Width="60px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Series No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_series_no" runat="server" Font-Size="11px" Width="60px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                            
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_qty" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="OS Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="rod_os_qty" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="rod_uom" runat="server" Font-Size="11px" Enabled="false" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Number Per UOM">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_pcs_per_uom" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <FooterTemplate>
                                   Total:
                                </FooterTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                                <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total Number" AccessibleHeaderText="rod_tot_pcs">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_tot_pcs" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <FooterTemplate>
                                   <asp:Label ID="rod_sub_total_pcs" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:Label>
                                </FooterTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                                <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Length">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_length" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Width">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_width" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Height">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_height" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Weight(KG)">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_kg" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="CBM">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_cbm" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
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
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" value="Back" onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=IB_RO_C <%else %>../../IMPORT/UL_RO/UploadRO.aspx<%end if %>'"
                            class="all_button" />   
                </td>
            </tr>
        </table>
    </div>
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
