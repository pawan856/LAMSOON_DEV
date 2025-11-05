<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PutAWay.aspx.vb" Inherits="PutAWay" ValidateRequest="false" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Put Away</title>
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

function LocLookUp(wh_val, lb_id, hd_id, wh_id) {
    if(document.piform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        if (wh_id)
            setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_id + "||WH");
        else
            setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
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

function maskNumOnly(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if (iKeyCode >= 48 && iKeyCode <= 57) return true;
    return false;
}

function resetLoc() {
    dsp_app_loc.innerHTML = "";
    document.piform.app_loc.value = "";

    if (document.piform.dtlSuffixList.value) {
        var i;
        var tempArray = (document.piform.dtlSuffixList.value).split(", ");

        for (i = 0; i < tempArray.length; i++) {
            if (eval("document.piform." + tempArray[i] + "_chk_loc").checked) {
                eval("document.piform." + tempArray[i] + "_chk_loc").checked = false;
            }
        }
    }
}

function appLoc() {
    if (document.piform.dtlSuffixList.value) {
        var i;
        var tempArray = (document.piform.dtlSuffixList.value).split(", ");

        for (i = 0; i < tempArray.length; i++) {
            if (eval("document.piform." + tempArray[i] + "_chk_loc").checked) {
                eval("document.piform." + tempArray[i] + "_gra_wh").value = document.piform.app_wh.value;
                eval("document.piform." + tempArray[i] + "_gra_loc").value = document.piform.app_loc.value;
                eval(tempArray[i] + "_dsp_gra_loc").innerHTML = dsp_app_loc.innerHTML;
                //alert(tempArray[i]);
            }
        }
    }
}

function selAllCB() {
    if (document.piform.dtlSuffixList.value) {
        var i;
        var tempArray = (document.piform.dtlSuffixList.value).split(", ");

        for (i = 0; i < tempArray.length; i++)
            eval("document.piform." + tempArray[i] + "_chk_loc").checked = true;
    }
}

/*
function setBodyHeightToContentHeight() {
    document.body.style.height = Math.max(document.documentElement.scrollHeight, document.body.scrollHeight) + "px";
}

setBodyHeightToContentHeight();
$addHandler(window, "resize", setBodyHeightToContentHeight);
*/
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:DisableDeleteButton();window.focus();">
<form name="piform" id="piform" runat="server">
<asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
<table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 500px">
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
            <td class="menuTD" align="left" nowrap>
                <!--<asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />-->
                <asp:Button ID="saveBtn1" runat="server" Text="OK" CssClass="all_button" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>
            <td align="right" width="200px" nowrap>
                <table cellpadding="0" cellspacing="0">
                    <tr>
                        <td nowrap><asp:DropDownList ID="app_wh" runat="server" Font-Size="11px" /></td>
                       <%-- <td nowrap><asp:Label ID="dsp_app_loc" width="100" runat="server" Font-Size="11px" /><asp:HiddenField ID="app_loc" runat="server" /></td>--%>
                      <%--  <td nowrap><asp:Image ID="Image_App_Loc_LookUp" onclick="LocLookUp(document.piform.app_wh.value, 'dsp_app_loc', 'app_loc', 'app_wh')" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" /></td>--%>
                        <td nowrap><input type="button" value="Clear" onclick="javascript:resetLoc();" class="all_button" /></td>
                        <td nowrap><input type="button" value="Select All" onclick="javascript:selAllCB();" class="all_button" /></td>
                        <td nowrap><input type="button" value="Apply" onclick="javascript:appLoc();" class="all_button" /></td>
                    </tr>
                </table>                                                                                               
            </td>
            <td class="menuTD" align="right" nowrap>
                <asp:Button ID="btnReset" Text="Reset" CssClass="all_button" runat="server" />
                <asp:Button ID="btnPrint" Text="Print" CssClass="all_button" runat="server" visible="false" />
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left">
                <Columns>
                    <asp:TemplateField HeaderText="Seq No.">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="gra_disp_seq" runat="server" Font-Size="11px" maxlength="10" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="PA List No.">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="GRA_PA_LIST_NO" runat="server" Font-Size="11px" maxlength="20" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Code">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="gra_itm_code" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Stock No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Name">
                        <ControlStyle Width="170px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="dsp_itm_name" width="80" MaxLength="100" runat="server" Font-Size="11px" BorderWidth="0" BackColor="Transparent" readonly="true" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Packy Key">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="gra_pack_key" runat="server" Font-Size="11px" style="text-align:center"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="gra_pallet_no" runat="server" Font-Size="11px" style="text-align:center"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">
                        <ItemStyle Width="60px" Wrap="false" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="GRA_BATCH_NO" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Expiry Date">
                        <ItemStyle Width="60px" Wrap="false" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="GRA_EXPIRY_DATE" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Suggested Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="gra_sug_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Put Away Qty">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="gra_pa_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="RCV Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:Label ID="rcv_qty" runat="server" Font-Size="11px" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="WH">
                        <ItemStyle Width="70px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:DropDownList ID="gra_wh" runat="server" Font-Size="11px" AutoPostBack="True" OnSelectedIndexChanged="gra_wh_SelectedIndexChanged"/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <%--<asp:TemplateField HeaderText="Location">
                        <ControlStyle />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                              <asp:dropdownlist runat="server" ID="gra_loc" Font-Size="11px"/>
                        <%--<table border="0" cellpadding ="0" cellspacing ="0" width="100%">
                        <tr>
                            <td align="left" nowrap>
                                <asp:CheckBox ID="chk_loc" runat="server" />
                                <asp:Label ID="dsp_gra_loc" width="130" runat="server" Font-Size="11px" />
                                <asp:HiddenField ID="gra_loc" runat="server" />
                            </td>
                            <td align="right">
                                <asp:ImageButton ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                            </td>
                        </tr>
                        </table>--
                        </ItemTemplate>
                    </asp:TemplateField>--%>

                    <asp:TemplateField HeaderText="Batch No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:ComboBox ID="gra_loc" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                    <asp:TemplateField HeaderText="Rejected Qty.">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="gra_rej_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Reject Reason">
                        <ControlStyle />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:DropDownList runat="server" ID="gra_rej_reason"  Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Final Location">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:ComboBox ID="gra_remark" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>

                    <%--<asp:TemplateField HeaderText="Remarks">
                        <ControlStyle />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="gra_remark" runat="server" Font-Size="11px" Width="130px" maxlength="500" />
                        </ItemTemplate>
                    </asp:TemplateField>--%>

                    <asp:TemplateField ControlStyle-Width="50px">
                        <ItemTemplate>
                            <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px"
                                CommandName="SplitItem" Text="Split" CssClass="all_button" Font-Bold="false" />
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
                <AlternatingRowStyle CssClass="REQUIRED" />
                <RowStyle CssClass="REQUIRED" />
                <EmptyDataRowStyle CssClass="REQUIRED" />
                <EditRowStyle CssClass="REQUIRED" />
            </asp:GridView>
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
<asp:HiddenField ID="IMP_CODE" runat="server" />
<asp:HiddenField ID="STORER_CODE" runat="server" />
<asp:HiddenField ID="GR_DOC_NO" runat="server" />
<asp:HiddenField ID="GR_DOC_TYPE" runat="server" />
<asp:HiddenField ID="editMode" runat="server" />
<asp:HiddenField ID="moduleAction" runat="server" />
<asp:HiddenField ID="dtlSuffixList" runat="server" />





    <asp:Panel ID="pnlResrvLoc" runat="server" CssClass="modalPopup" style="display: none">
    <table width="800px" cellpadding="0" cellspacing="0" border="0">
    <tr style="height:25px">
        <td class="TITLE">
            <asp:Panel runat="Server" ID="pnlResrvLocDrag" Style="cursor: move;">
            <table cellpadding="0" cellspacing="0" border="0" width="100%">
            <tr>
                <td style="text-align:left; width:50%" class="TITLE">
                    <font size="2">Reserved Location</font>
                </td>
                <td style="text-align:right" class="TITLE">
                    <asp:Button runat="server" ID="btnSelOtherLoc" Text="Select Other Location" CssClass="all_button" />
                </td>
            </tr>
            </table>
            </asp:Panel>
        </td>
    </tr>
    <tr>
        <td>
            <asp:UpdatePanel runat="server" ID="updtPnl_ResrvLocList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
                <asp:GridView ID="gvResrvLocList" runat="server" Width="100%" Font-Names="Arial"
                    Font-Overline="False" Font-Size="10px" 
                    AutoGenerateColumns="False" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" ShowHeader="True"
                    CellPadding="2" CaptionAlign="Top" GridLines="None" CellSpacing="1" ShowHeaderWhenEmpty="true" 
                    HorizontalAlign="Left">
                    <RowStyle CssClass="GV" />
                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                    <Columns>
                    <asp:boundfield datafield="LOBH_CODE"  headertext="Reserv. No." HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_LOC"  headertext="Location Code" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_SKU_NO"  headertext="Stock No." HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="PACK_KEY"  headertext="Pack Key" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_ITM_DESC"  headertext="Item Name" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_CBM"  headertext="CBM" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_CBM_PERC"  headertext="%" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:boundfield datafield="LOBD_STATUS"  headertext="Status" HeaderStyle-Font-Bold="false" ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" HeaderStyle-Wrap="false"/>
                    <asp:TemplateField>
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" Width="100px" />
                        <ItemTemplate>
                            <asp:Button ID="btnResrvLocSelect" Text="Select" CssClass="all_button" CommandName="SEL_LOC" runat="server" />
                            <asp:HiddenField ID="LOBD_LOC_HDN" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    </Columns>
                </asp:GridView>
                <asp:HiddenField ID="tarLocID" runat="server" />
                <asp:HiddenField ID="tarDspLocID" runat="server" />
                <asp:HiddenField ID="tarWhID" runat="server" />
            </ContentTemplate>
            </asp:UpdatePanel>
        </td>
    </tr>
    <tr>
        <td class="TITLE" align="left">
            <asp:Button runat="server" ID="btnCloseResrvLoc" Text="Close" CssClass="all_button" />
        </td>
    </tr>
    </table>
    </asp:Panel>

    <asp:ModalPopupExtender ID="resrvLoc_ModalPopupExtender" runat="server"
        DynamicServicePath="" 
        Enabled="True" 
        TargetControlID="resrvLoc_dummy" 
        PopupControlID="pnlResrvLoc"
        BackgroundCssClass="modalBackground"
        DropShadow="true"         
        CancelControlID="btnCloseResrvLoc"
        PopupDragHandleControlID="pnlResrvLocDrag" 
        RepositionMode="None" 
        BehaviorID="resrvLoc_behavior"
        Drag="true"
        Y="50"
        />

    <asp:HiddenField ID="resrvLoc_dummy" runat="server" /> 



    <!-- Update panel for asyn postback -->
    <asp:UpdatePanel runat="server" ID="updtPnlPostBack" UpdateMode="Conditional">
    <ContentTemplate />
    </asp:UpdatePanel>

    <!-- Update panel that decided for validation alert -->
    <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>










</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
