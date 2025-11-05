<%@ Page Language="VB" AutoEventWireup="false" CodeFile="COMain_C.aspx.vb" Inherits="OUTBOUND_COMain_C" MaintainScrollPositionOnPostback="true" %>
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
    <div runat="server" id="DIVSCRIPT">
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

    function ItemLookUpOld(STORER_CODE) {
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
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "ItemLookUp";
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

            window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");
            
            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM_BAL");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemKeyList|1, locKeyList|2, qtyList|CC_ITEM_QTY#txt_item_qty");
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
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "sbLookUp";
            document.hiddenForm.submit();
        }
    }

    function sumqtytotal(qty) {
        var DataGridObj = document.getElementById('GridView1');

        var totAmt = 0;
        var eachQty;

        if (DataGridObj.rows.length > 0) {
            for (var iRow = 1; iRow <= DataGridObj.rows.length - 1; iRow++) {
                if (iRow == DataGridObj.rows.length -1) {
                    DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].innerText = totAmt;
                    DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].align = "Right";
                    DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].style.fontWeight = "bold";
                } else {
                    eachQty = 0

                    if (DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].children[0].value != "") {
                        eachQty = parseFloat(DataGridObj.rows[iRow].cells[<%=ViewState("qtyIndex")%>].children[0].value.replace(/,/gi, ""));
                    }
                        totAmt = parseFloat(totAmt) + parseFloat(eachQty);
                    }
            }
        } 
    }
    
    function HoldStock(storer_code, co_code, cod_seq, qty_id, batch_id) {
        if(document.myform.editMode.value!="V")
        {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "holdMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=768,left=5,top=15,resizable=yes");
        
            setInterfaceDataToForm(document.hiddenForm, "storer_code", storer_code);
            setInterfaceDataToForm(document.hiddenForm, "co_code", co_code);
            setInterfaceDataToForm(document.hiddenForm, "cod_seq", cod_seq);
            setInterfaceDataToForm(document.hiddenForm, "qty_id", qty_id);
            setInterfaceDataToForm(document.hiddenForm, "batch_no", document.getElementById(batch_id).value);
        
            document.hiddenForm.action = "COHMain.aspx";
            document.hiddenForm.target = "holdMain";
            document.hiddenForm.submit();
        }
    }

    function updtVend(vendCtrl) {
        var itmCtrlID = (vendCtrl.id).replace("cod_vnd_code", "h_cod_itm_code");
        var packKeyCtrlID = (vendCtrl.id).replace("cod_vnd_code", "h_cod_pack_key");

        if (vendCtrl.value != "")
            PageMethods.updateVendCode(document.getElementById("STORER_CODE").value, document.getElementById(itmCtrlID).value, document.getElementById(packKeyCtrlID).value,  vendCtrl.id, vendCtrl.value, updtVendReturn);
    }

    function updtVendReturn(returnVendInfo) {
//        if (returnVendInfo.PCS_UOM != "")
//            document.getElementById(returnVendInfo.PCS_UOM_ID).value = returnVendInfo.PCS_UOM;

        if (returnVendInfo.VOL != "")
            document.getElementById(returnVendInfo.VOL_ID).value = returnVendInfo.VOL;

        if (returnVendInfo.WGT != "")
            document.getElementById(returnVendInfo.WGT_UOM_ID).value = returnVendInfo.WGT;
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
    <asp:HiddenField ID="vndList" runat ="server" />    

    <asp:HiddenField ID="itemKeyList" runat ="server" />
    <asp:HiddenField ID="locKeyList" runat ="server" />
    <asp:HiddenField ID="qtyList" runat ="server" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
            <tr>
                <td colspan="6" class="TITLE">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td class="TITLE" align="left">
                            Customer Order:
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
                        <input id="btnBack2" type="button" value="Back" onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=OB_CO_C <%else %>../../IMPORT/UL_CO/UploadCO.aspx<%end if %>'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" />
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td colspan="6" class="TITLE">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:Label ID="CO_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="3">
                        <asp:Label ID="CO_STATUS" runat="server" /></font>
                        <asp:HiddenField runat="server" ID="STORER_CODE" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_CUS_REF_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="3">
                        <asp:TextBox ID="CO_CUS_REF_NO" runat="server" MaxLength="20" Width="130px"></asp:TextBox>
                    </font>
                </td>
                
            </tr>                 
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label runat="server" ID="lbl_CO_CUST_INV_NO" text="Customer Invoice No." />
                    </font>
                </td>
                <td colspan="5">
                    <asp:TextBox runat="server" ID="CO_CUST_INV_NO" MaxLength="50" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_DATE" runat="server" /></font>
                </td>
                <td width="150px"  colspan="5">
                    <asp:TextBox ID="CO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="CO_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>               
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CUS_CODE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                    <td width="50px" runat="server"><asp:DropDownList ID="CUS_CODE" runat="server" MaxLength="20" 
                        AutoPostBack="True"></asp:DropDownList></td>
                    <td><font size="3"><asp:textbox ID="CUS_NAME" runat="server" ReadOnly="true" BackColor="#dddddd" Width="365px" /></font></td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <asp:Label ID="lbl_co_addr" runat="server" />
                </td>
                <td colspan="5">
                        <asp:TextBox ID="CO_ADDR1" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                    <br />
                        <asp:TextBox ID="CO_ADDR2" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                    <br />
                        <asp:TextBox ID="CO_ADDR3" runat="server" MaxLength="100" Width="494px"></asp:TextBox>
                 
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_AREA_DEL" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:TextBox ID="CO_AREA_DEL" runat="server" MaxLength="50" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_REGION_DEL" runat="server" /></font>
                </td>
                <td>
                    <font size="3">
                        <asp:TextBox ID="CO_REGION_DEL" runat="server" MaxLength="50" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_COUNTRY_DEL" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="CO_COUNTRY_DEL" runat="server" Width ="130px" MaxLength="50"></asp:TextBox>
                </td>                
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_co_cus_cont" runat="server" /></font>
                </td>
                <td colspan="5">
                    <font size="3">
                        <asp:TextBox ID="CO_CUS_CONT" runat="server" MaxLength="50" Width="165px"></asp:TextBox>
                    </font>
                </td>
            </tr>          
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_co_cus_cont_tel" runat="server" /></font>
                </td>
                <td colspan="5">
                    <font size="3">
                        <asp:TextBox ID="CO_CUS_CONT_TEL" runat="server" MaxLength="20" 
                        Width="165px" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
            </tr>                        
            <tr>
                <td class="LabelTD" nowrap valign="top">
                    <font size="3">
                        <asp:Label ID="lbl_CO_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="CO_REM" runat="server" Height="117px" Width="422px" 
                        TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>
            </tr>
            <tr>
            <!--CO_TARGET_DELDATE-->
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_TARGET_DELDATE" runat="server" Text="Target Delivery Date" /></font>    
                </td>
                <td >
                    <asp:TextBox ID="CO_TARGET_DELDATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="BTN_CO_TARGET_DELDATE" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="CO_TARGET_DELDATE" PopupButtonID="BTN_CO_TARGET_DELDATE" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="3">
                        <asp:Label ID="lbl_CO_SHIP_MODE" runat="server" TEXT="Ship Mode" /></font>    
                </td>
                <td colspan="3">
                    <asp:DropDownList runat="server" ID="CO_SHIP_MODE" />
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
        <asp:panel runat="server" ID="unusedFileds" Visible = "false">
            <asp:Label ID="lbl_CO_INV_NO" runat="server" />
            <asp:TextBox ID="CO_INV_NO" runat="server" MaxLength="20" Width="130px"></asp:TextBox>
            <asp:Label ID="lbl_CO_TRACK_NO" runat="server" />
            <asp:TextBox ID="CO_TRACK_NO" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
            <asp:Label ID="lbl_CO_FTRACK_NO" runat="server" />\
            <asp:TextBox ID="CO_FTRACK_NO" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
            <asp:Label ID="lbl_PRJ_CODE" runat="server" />
            <asp:DropDownList ID="CO_PROJECT_NO" runat="server" MaxLength="20" AutoPostBack="True"></asp:DropDownList>
            <asp:textbox ID="PRJ_NAME" runat="server" ReadOnly="true" BackColor="#dddddd" Width="365px" />
        </asp:panel>

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
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" 
                        CssClass="all_button" UseSubmitBehavior ="false" Width="96px" />                    
                    <asp:Button ID="btnExportTF" runat="server" Text="Export Data File" CssClass="all_button" UseSubmitBehavior="false" visible="false"/>                        
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="COD_SEQ" HorizontalAlign="Left" ShowFooter="true">
                           <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:TemplateField HeaderText="Seq.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_disp_seq" runat="server" Font-Size="11px" style="text-align:center" Width="30px" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="30px">
                                <ItemTemplate>
                                    <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="SplitItem" Text="S" CssClass="all_button" Font-Bold="false" UseSubmitBehavior="false" />
                                </ItemTemplate>
                                <ControlStyle Width="30px"></ControlStyle>
                                <HeaderStyle Width="30px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Plan">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_plant" runat="server" Font-Size="11px" Width="50px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Ref. No.">
                                <ControlStyle Width="90px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_ref_no" runat="server" Font-Size="11px" Width="90px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                            
                             <asp:TemplateField HeaderText="Carton No.">
                                <ControlStyle Width="90px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>                                    
                                    <asp:TextBox ID="cod_carton_no" runat="server" Font-Size="11px" Width="90px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="90px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>                                    
                                    <asp:TextBox ID="cod_pallet_no" runat="server" Font-Size="11px" Width="90px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>                            
                            <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="cod_batch_no" />
                                    <asp:HiddenField runat="server" ID="ori_batch" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Lot No.">
                                <ControlStyle Width="90px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_lot_no" runat="server" Font-Size="11px" Width="90px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Vendor Code">
                                <ControlStyle Width="120px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropdownList ID="cod_vnd_code" runat="server" Font-Size="11px" Width="120px" onchange="javascript:updtVend(this);"></asp:DropdownList>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Code">
                                <ControlStyle Width="120px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="cod_itm_code" runat="server" Font-Size="11px" Width="120px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Stock No.">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="180px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="cod_itm_desc" runat="server" Font-Size="11px" Width="180px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Description">
                                <ControlStyle Width="180px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_desc" runat="server" Font-Size="11px" Width="180px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Expiry Date">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="dsp_COD_EXPIRY_DATE" />
                                    <asp:HiddenField runat="server" ID="COD_EXPIRY_DATE" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Manu. Date">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" ID="dsp_COD_MANU_DATE" />
                                    <asp:hiddenfield runat="server" ID="COD_MANU_DATE" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Type">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_packing" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <FooterTemplate>
                                   Total Qty:
                                </FooterTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                                <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="cod_pack_key" runat="server" Font-Size="11px" Width="60px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty" AccessibleHeaderText="cod_qty">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_qty" runat="server" Font-Size="11px" Width="60px" style="text-align:right" onkeypress="return maskKey(event);" Text=""></asp:TextBox>
                                </ItemTemplate>
                                <FooterTemplate>
                                   <asp:Label ID="cod_total_qty" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                </FooterTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                                <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="OS Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="cod_os_qty" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty/ Carton">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_pcs_carton" runat="server" Font-Size="11px" 
                                        style="text-align:right" Width="60px" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Carton/ Pallet">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_carton_pallet" runat="server" Font-Size="11px" 
                                        style="text-align:right" Width="60px" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>                            
                            <asp:TemplateField HeaderText="UOM">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="cod_uom" runat="server" Font-Size="11px" Enabled="false" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                              <asp:TemplateField HeaderText="Number/ UOM">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_pcs_uom" runat="server" Font-Size="11px" 
                                        style="text-align:right" Width="60px" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total Number">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_totpcs" runat="server" Font-Size="11px" Width="60px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Weight(kg)">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_tot_wgt" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="CBM">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_tot_cbm" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="130px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="cod_rem" runat="server" Width="130px" Font-Size="10px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                    <asp:HiddenField runat="server" ID="h_cod_itm_code" />
                                    <asp:HiddenField runat="server" ID="h_cod_pack_key" />
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
                    <input id="btnBack" type="button" value="Back" onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=OB_CO_C <%else %>../../IMPORT/UL_CO/UploadCO.aspx<%end if %>'"
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