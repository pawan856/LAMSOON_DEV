<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PickList.aspx.vb" Inherits="PickList" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Pick List</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language="javascript" type="text/javascript" src="../../js/validation.js"></script>
<script language="javascript" type="text/javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" type="text/javascript" src="../../js/listUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" type="text/javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript" type="text/javascript">

function DisableDeleteButton() {
    var grp = document.getElementsByName("btnDelete");
    var count;
    count = grp.length;
    if (count == 1) {
       grp[0].disabled = true;
    }
}

function LocLookUp_old(wh_val, lb_id, hd_id, wh_id, lb_fl, hd_fl, lb_ar, hd_ar, lb_rk, hd_rk, lb_bn, hd_bn) {
    if(document.piform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_id + "||WH, " + lb_fl + "|L|FL, " + hd_fl + "||FL, " + lb_ar + "|L|AR, " + hd_ar + "||AR, " + lb_rk + "|L|RK, " + hd_rk + "||RK, " + lb_bn + "|L|BN, " + hd_bn + "||BN");
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function LocLookUp(itm_code, pack_key, pallet_no, batch_no, sku_no, orgLoc, dspSeq) {
    if(document.piform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        setInterfaceDataToForm(document.hiddenForm, "IMP_CODE", document.piform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.piform.STORER_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "ITM_CODE", itm_code);
        setInterfaceDataToForm(document.hiddenForm, "ITM_SKU_NO", sku_no);
        setInterfaceDataToForm(document.hiddenForm, "PACK_KEY", pack_key);
        setInterfaceDataToForm(document.hiddenForm, "PALLET_NO", pallet_no);
        //setInterfaceDataToForm(document.hiddenForm, "ref_no", opener.myform.DO_CUS_REF_NO.value);
        setInterfaceDataToForm(document.hiddenForm, "BATCH_NO", batch_no);
        setInterfaceDataToForm(document.hiddenForm, "ORG_LOC", orgLoc);

        setInterfaceDataToForm(document.hiddenForm, "DOD_DISP_SEQ", dspSeq);

        setInterfaceDataToForm(document.hiddenForm, "CO_CODE", document.piform.CO_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_CODE", document.piform.DO_CODE.value);

        document.hiddenForm.action = "./PickListLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function LocSerialLookUp(itm_code, pack_key, pallet_no, batch_no, sku_no, qty2, seq_no, orgLoc, dspSeq) {
    if (document.piform.editMode.value != "V") {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");

        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        setInterfaceDataToForm(document.hiddenForm, "IMP_CODE", document.piform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.piform.STORER_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "ITM_CODE", itm_code);
        setInterfaceDataToForm(document.hiddenForm, "ITM_SKU_NO", sku_no);
        setInterfaceDataToForm(document.hiddenForm, "PACK_KEY", pack_key);
        setInterfaceDataToForm(document.hiddenForm, "PALLET_NO", pallet_no);
        //setInterfaceDataToForm(document.hiddenForm, "ref_no", opener.myform.DO_CUS_REF_NO.value);
        setInterfaceDataToForm(document.hiddenForm, "BATCH_NO", batch_no);
        setInterfaceDataToForm(document.hiddenForm, "QTY2", qty2);
        setInterfaceDataToForm(document.hiddenForm, "PLD_SEQ", seq_no);
        setInterfaceDataToForm(document.hiddenForm, "ORG_LOC", orgLoc);

        setInterfaceDataToForm(document.hiddenForm, "DOD_DISP_SEQ", dspSeq);

        setInterfaceDataToForm(document.hiddenForm, "CO_CODE", document.piform.CO_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_CODE", document.piform.DO_CODE.value);

        //document.hiddenForm.action = "./PickSerialLookup.aspx";
        document.hiddenForm.action = "./PickSerialLookupSplit.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function saveGVDT() {
    document.piform.moduleAction.value = "SAVEGVDT";
    //document.piform.submit();

    __doPostBack("updtPnlPostBack", "");
}

function refreshGV()
{
    document.piform.moduleAction.value = "RELOADPL";
    document.piform.submit();
}

function saveok()
{
    document.piform.moduleAction.value = "SAVEOK";
    document.piform.submit();
    }

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:DisableDeleteButton();">
<form name="piform" id="piform" runat="server">
<asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 600px">
    <tr>
        <td class="TITLE" colspan="8">
            <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="100%" class="TITLE">
                        <b>
                            <asp:Label ID="lheader" runat="server" /></b>
                    </td>
                    <td width="100%" class="TITLE">
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="menuTD" align="left">
                <!--<asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />-->
                <asp:Button ID="saveBtn1" runat="server" Text="OK" CssClass="all_button" />
                <asp:Button ID="btnRefresh" Text="Refresh Stock Qty" CssClass="all_button" Visible="false" runat="server" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td class="menuTD" align="right">
                <asp:Button ID="btnReset" Text="Reset" CssClass="all_button" runat="server" />
                <asp:Button ID="btnPrint" Text="Print" CssClass="all_button" runat="server" visible="false" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:UpdatePanel runat="server" ID="updtPnl_gvPickList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left" >
                <Columns>
                    <asp:TemplateField HeaderText="Seq No.">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dod_disp_seq" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_seq" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>                        
                    <asp:TemplateField HeaderText="Picked By">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_picked_by" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Code">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="pld_item_no" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Stock No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="DO Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="pld_do_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Location FFI Qty">
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
                    <asp:TemplateField HeaderText="EDI Code">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="bn_csms_code" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="WH">
                        <ItemStyle Width="70px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:DropDownList ID="pld_wh" runat="server" Font-Size="11px" AutoPostBack="True" OnSelectedIndexChanged="pld_wh_SelectedIndexChanged"/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Org Loc">
                        <ControlStyle Width="110px" />
                        <HeaderStyle />
                        <ItemStyle />
                        <ItemTemplate>
                            <asp:Label ID="pld_org_loc" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Location">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:ComboBox ID="pld_loc" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>

                    <%--<asp:TemplateField HeaderText="Location">
                        <ControlStyle />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                             <asp:dropdownlist runat="server" ID="pld_loc" Font-Size="11px"/>
                        <table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                        <tr>
                            <td align="left">
                                <asp:Label ID="dsp_pld_loc" width="100px" runat="server" Font-Size="11px" />
                                <asp:HiddenField ID="pld_loc" runat="server" />
                            </td>
                            <td align="right">
                                <asp:Imagebutton ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                            </td>
                        </tr>
                        </table>
                        </ItemTemplate>
                    </asp:TemplateField>--%>
                    <asp:TemplateField HeaderText="Floor">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dsp_pld_floor" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_floor" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Area">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dsp_pld_area" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_area" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Rack">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dsp_pld_rack" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_rack" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField> 
                    <asp:TemplateField HeaderText="Bin">
                        <ControlStyle Width="25px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="dsp_pld_bin" width="70" runat="server" Font-Size="11px" style="text-align:center" />
                            <asp:HiddenField ID="pld_bin" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Location Available Qty">
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
                    <asp:TemplateField HeaderText="Hold Qty / Total Bal">
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
                    <asp:TemplateField HeaderText="Expiry Date">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="iloc_expiry_date" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty 2">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_qty2" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Serial">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_serial_no" runat="server" Font-Size="11px" maxlength="80" />
                             <asp:Button ID="btnRedrum" runat="server" CommandName="REDRUM" CssClass="all_button" visible="false" Text="Re-Drum" />
                             <asp:HiddenField ID="from_drum_id" runat="server" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemStyle Width="80px" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_batch_no" runat="server" Font-Size="11px" maxlength="80" />

                            <%--<asp:Label ID="pld_batch_no" runat="server" Font-Size="11px" maxlength="20" />--%>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Packy Key">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="pld_pack_key" runat="server" Font-Size="11px" style="text-align:center"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="55px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_pallet_no" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Short ship Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="PLD_SS_QTY" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>             

                    <asp:TemplateField HeaderText="Remarks">
                        <ControlStyle Width="150px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="pld_remark" runat="server" Font-Size="11px" Width="150px" maxlength="200" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:TemplateField>
                    <asp:TemplateField ControlStyle-Width="50px">
                        <ItemTemplate>
                            <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px" Visible="false"
                                CommandName="SplitItem" Text="Split" CssClass="all_button" Font-Bold="false" />
                        </ItemTemplate>
                        <ControlStyle Width="50px"></ControlStyle>
                        <HeaderStyle Width="50px" />
                    </asp:TemplateField>
                    <asp:TemplateField ControlStyle-Width="50px">
                        <ItemTemplate>
                            <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px" Visible="false"
                                CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                        </ItemTemplate>
                        <ControlStyle Width="50px"></ControlStyle>
                        <HeaderStyle Width="50px" />
                    </asp:TemplateField>
                </Columns>
                <AlternatingRowStyle CssClass="REQUIRED" />
                <RowStyle CssClass="REQUIRED" />
                <EmptyDataRowStyle CssClass="REQUIRED" />
                <EditRowStyle CssClass="REQUIRED" />
            </asp:GridView>
            </ContentTemplate>
            </asp:UpdatePanel>
        </td>
    </tr>
    <tr>
        <td colspan="8" class="menuTD">
            <asp:Button ID="saveBtn2" runat="server" Text="OK" CssClass="all_button" />
            <input id="btnClose2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                class="all_button" />
        </td>
    </tr>
</table>

    <!-- Update panel for asyn postback -->
    <asp:UpdatePanel runat="server" ID="updtPnlPostBack" UpdateMode="Conditional">
    <ContentTemplate />
    </asp:UpdatePanel>

    <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>

    <asp:UpdatePanel runat="server" ID="updtPnlKey">
    <ContentTemplate>
        <asp:HiddenField ID="editMode" runat="server" />
        <asp:HiddenField ID="moduleAction" runat="server" />
        <asp:HiddenField ID="hd_drumList" runat="server" />
    </ContentTemplate>
    </asp:UpdatePanel>
    
    <asp:UpdatePanel runat="server" ID="DrumUDP" RenderMode="Inline">
        <ContentTemplate>
            <asp:Panel runat="server" CssClass="modalPopup" ID="drumPanel" Width="800px" Style="display: none"> <!---->
                <table style="border:0; width:100%;"  >
                     <tr>
                        <td colspan="3" align="left" class="LabelTD">Cable Re-Drum</td>
                    </tr>
                    <tr>
                        <td align="left">From Drum ID:&nbsp;&nbsp;&nbsp;&nbsp;<asp:label runat="server" ID="FromDrumID_TXT" /></td>
                        <td></td>
                        <td align="left">To Drum ID: <asp:textbox runat="server" ID="ToDrumID_TXT" MaxLength="20" /></td>
                    </tr>
                    <tr>
                        <td colspan="3" align="center"><asp:Button runat="server" ID="btnConfirmP" Text="Confirm" CssClass="all_button" />&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button runat="server" ID="btnCancelP" Text="Cancel" CssClass="all_button" /></td>
                    </tr>
                    <tr>
                        <td colspan="3" align="center">&nbsp;</td>
                    </tr>
                    <tr>
                        <td valign="top">
                              <asp:GridView ID="GVDrumIN" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="ILOC_SEQ, ILBS_SEQ" HorizontalAlign="Left">
                              <RowStyle CssClass="GV" />
                                <Columns>
                                      <asp:TemplateField HeaderText="Serial Number">
                                        <ControlStyle Width="110px" />
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_SERIAL_NO" runat="server" Font-Size="11px" />
                                            <asp:hiddenfield ID="ILOC_SEQ" runat="server" />
                                            <asp:hiddenfield ID="ILBS_SEQ" runat="server" />
                                            <asp:hiddenfield ID="dFlag" runat="server" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="Drum Lv">
                                        <ControlStyle Width="110px" />
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_DRUM_LEVEL" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="Qty.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_QTY2" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="right" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="UOM">
                                      <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_UOM2" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                      <asp:TemplateField>
                                      <HeaderStyle HorizontalAlign="Left" Width="20px" />
                                        <ItemTemplate>
                                            <asp:CheckBox runat="server" ID="selectYN" visible="false"/>
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                </Columns>
                              <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                            </asp:GridView>
                        </td>
                        <td><asp:Button runat="server" ID="btnMove" Text=">>" Font-Bold="true" /></td>
                        <td valign="top">
                            <asp:GridView ID="GVDrumOUT" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" ShowHeaderWhenEmpty="true" 
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="ILOC_SEQ, ILBS_SEQ" HorizontalAlign="Left">
                              <RowStyle CssClass="GV" />
                                <Columns>
                                      <asp:TemplateField HeaderText="Serial Number">
                                        <ControlStyle Width="110px" />
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_SERIAL_NO" runat="server" Font-Size="11px" />
                                            <asp:hiddenfield ID="ILOC_SEQ" runat="server" />
                                            <asp:hiddenfield ID="ILBS_SEQ" runat="server" />
                                            <asp:hiddenfield ID="dFlag" runat="server" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="Drum Lv">
                                        <ControlStyle />
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:textbox ID="ILBS_DRUM_LEVEL" runat="server" Font-Size="11px" MaxLength="20" width="50px" />
                                            <asp:FilteredTextBoxExtender ID="FilteredTextBoxExtender1" runat="server" TargetControlID="ILBS_DRUM_LEVEL" FilterType="Numbers" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="Qty.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:label ID="ILBS_QTY2" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="right" />
                                      </asp:TemplateField>
                                      <asp:TemplateField HeaderText="UOM">
                                      <HeaderStyle HorizontalAlign="Left" />
                                      <ItemTemplate>
                                            <asp:label ID="ILBS_UOM2" runat="server" Font-Size="11px" />
                                      </ItemTemplate>
                                      <ItemStyle HorizontalAlign="Left" />
                                      </asp:TemplateField>                                                                          
                                </Columns>
                              <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                            </asp:GridView>                
                        </td>
                    </tr>            
                </table>
                <asp:HiddenField runat="server" id="FromDrum" />
                <asp:HiddenField runat="server" id="ToDrum_HD" />
                <asp:HiddenField runat="server" id="ToIloc_seq" />
                <asp:HiddenField runat="server" id="selectedSeqList" />
                <asp:HiddenField runat="server" id="FromRowIDX" />
            </asp:Panel>
            <asp:ModalPopupExtender ID="drumPop" runat="server" 
            DynamicServicePath="" 
            Enabled="True" 
            TargetControlID="dummy1" 
            PopupControlID="drumPanel"
            BackgroundCssClass="modalBackground"
            DropShadow="true"         
            CancelControlID="btnCancelP"
            Y="50">
        
            </asp:ModalPopupExtender>
            <asp:HiddenField ID="dummy1" runat="server" />
        </ContentTemplate>
    </asp:UpdatePanel>


    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="STORER_CODE" runat="server" />
    <asp:HiddenField ID="CO_CODE" runat="server" />
    <asp:HiddenField ID="DO_CODE" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
