<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SADJMain.aspx.vb" Inherits="OPERATION_STA_STAMain" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Stock Adjustment</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

    <script language="javascript" src="../../js/validation.js"></script>
    <script language="javascript" src="../../js/JS_Calendar.js"></script>
    <script language="javascript" src="../../js/listUtil.js"></script>
    <script language="javascript" src="../../js/formatUtil.js"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js"></script>
    <script src="../../js/jquery-1.9.1.min.js"></script>
    <script src="../../js/jquery-migrate-1.1.1.min.js"></script>

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

            if (parseFloat(oQty) > parseFloat(rQty)) {
                vQty = (parseFloat(oQty) - parseFloat(rQty));
                eval("document.getElementById('" + lvqid + "')").innerHTML = "-" + vQty;
                eval("document.getElementById('" + vqid + "')").value = "-" + vQty;
            } else if (parseFloat(oQty) < parseFloat(rQty)) {
                vQty = (parseFloat(rQty) - parseFloat(oQty));
                eval("document.getElementById('" + lvqid + "')").innerHTML = vQty;
                eval("document.getElementById('" + vqid + "')").value = vQty;
            } else if (parseFloat(oQty) == parseFloat(rQty)) {
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

        function getLoad() {
            var load_modalPopup = $find('load_ModalPopupExtender');
            load_modalPopup.show();
        }

        $(document).keypress(
            function (event) {
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
        <input type="hidden" name="moduleAction" value="" />
        <asp:HiddenField ID="itemList" runat="server" />
        <asp:HiddenField ID="packKeyList" runat="server" />
        <asp:HiddenField ID="seqList" runat="server" />
        <div id="div1">
            <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
                <tr>
                    <td colspan="4" class="TITLE" align="left">
                        <table border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="100%" class="TITLE">
                                    <b>
                                        <asp:Label ID="lheader" runat="server" /></b>
                                </td>
                                <td width="100%" class="TITLE">
                                    <asp:Table ID="cmBar" runat="server" border="0" CellSpacing="0" CellPadding="0"></asp:Table>
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
                                    
                                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                                        class="all_button" />
                                </td>
                                <td class="menuTD" align="right">
                                    <%If Session("pagemode") <> "N" Then%>
                                    <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" />                                    
                                    <% End If%>
                                     <asp:Button ID="CancelBtn" Text="Cancel" CssClass="all_button" runat="server" />
                                   
                                </td>
                            </tr>
                        </table>

                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_CODE" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:Label ID="AD_CODE" runat="server" />
                        </font>
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_STATUS" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:Label ID="AD_STATUS" runat="server" /></font>
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
                            <asp:Label ID="lbl_AD_TYPE" runat="server" /></font>
                    </td>
                    <td runat="server">
                        <font size="2">
                            <asp:DropDownList ID="AD_TYPE" runat="server"></asp:DropDownList>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_DATE" runat="server" /></font>
                    </td>
                    <td width="150px">
                        <asp:TextBox ID="AD_DATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                        <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif"
                            ImageAlign="Middle" BorderWidth="0" Style="padding-left: 4px" />
                        <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar"
                            TargetControlID="AD_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_BY" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="AD_BY" runat="server" MaxLength="20"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_REF_NO" runat="server" /></font>
                    </td>
                    <td nowrap colspan="3">
                        <font size="2">
                            <asp:TextBox ID="AD_REF_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_REF_CK_CODE" runat="server" /></font>
                    </td>
                    <td nowrap colspan="3">
                        <font size="2">
                            <asp:Label ID="AD_REF_CK_CODE" runat="server" />
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_WH" runat="server" /></font>
                    </td>
                    <td colspan="3" runat="server">
                        <font size="2">
                            <asp:DropDownList ID="AD_WH" runat="server"></asp:DropDownList>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_AD_REM" runat="server" /></font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="AD_REM" runat="server" Height="117px" Width="523px"
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
                                    <b>
                                        <asp:Label ID="lbl_ImageHd" runat="server" /></b>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="menuTD">
                        <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible="false" />
                        <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" />
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                            Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                            CellPadding="3" CaptionAlign="Top" DataKeyNames="ad_seq" HorizontalAlign="Left">
                            <RowStyle CssClass="GV" />
                            <Columns>
                                <asp:BoundField DataField="ad_seq" HeaderText="No.">
                                    <ControlStyle Width="30px" />
                                    <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Item Code">
                                    <ControlStyle Width="100px" />
                                    <ItemTemplate>
                                        <asp:HiddenField ID="add_itm_code" runat="server" />
                                        <asp:HiddenField ID="add_pallet_no" runat="server" />
                                        <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" Width="100px"></asp:Label>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pack Key">
                                    <ControlStyle Width="60px" />
                                    <ItemTemplate>
                                        <asp:Label ID="add_pack_key" runat="server" Font-Size="11px" Width="60px"></asp:Label>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Item Name">
                                    <ItemTemplate>
                                        <asp:Label ID="itm_name" runat="server" Font-Size="11px" Width="120px"></asp:Label>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Lot">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="disp_add_batch_no" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="add_batch_no" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Loc.">
                                    <ItemTemplate>
                                        <asp:TextBox ID="add_loc" runat="server" Font-Size="11px" Width="60px" Style="text-align: right"></asp:TextBox>
                                        <%--<asp:Label ID="dsp_add_loc" Width="100px" runat="server" Font-Size="11px" CssClass="REQUIRED" />
                                        <asp:HiddenField ID="add_loc" runat="server" />--%>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Orig.<br>Qty">
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_add_org_qty" runat="server" Font-Size="11px" Width="60px" Style="text-align: right"></asp:Label>
                                        <asp:HiddenField ID="add_org_qty" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Revised<br>Qty">
                                    <ItemTemplate>
                                        <asp:TextBox ID="add_rev_qty" runat="server" Font-Size="11px" Width="60px" Style="text-align: right"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Variance<br>Qty">
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_add_var_qty" runat="server" Font-Size="11px" Width="60px" Style="text-align: right"></asp:Label>
                                        <asp:HiddenField ID="add_var_qty" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Remarks">
                                    <ControlStyle Width="110px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="add_rem" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Revise Expiry Date">
                                    <ItemTemplate>
                                        <asp:TextBox ID="add_manu_date" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                                        <asp:ImageButton ID="btnMDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif"
                                            ImageAlign="Middle" BorderWidth="0" Style="padding-left: 4px" Height="20px" Width="20px" />
                                        <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar"
                                            TargetControlID="add_manu_date" PopupButtonID="btnMDATE_ID1" Format="dd/MM/yyyy" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" Wrap="false" />
                                </asp:TemplateField>
                                <%-- Add new field--%>
                                <asp:TemplateField HeaderText="Drum Id">
                                    <ControlStyle Width="110px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="DRUM_ID" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Drum Level">
                                    <ControlStyle Width="110px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="DRUM_LEVEL" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <%-- end--%>

                                <asp:TemplateField HeaderText="Serial No.">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_ADD_SERIAL_NO" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="ADD_SERIAL_NO" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>

                                <%-- Add new field--%>
                                <asp:TemplateField HeaderText="UOM 2">
                                    <ControlStyle Width="110px" />
                                    <ItemTemplate>
                                        <asp:Label ID="ITM_UOM" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:Label>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <%-- end--%>


                                <asp:TemplateField HeaderText="Org. Qty.2">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_ADD_ORG_QTY2" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="ADD_ORG_QTY2" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Revised<br>Qty2">
                                    <ItemTemplate>
                                        <asp:TextBox ID="add_rev_qty2" runat="server" Font-Size="11px" Width="60px" Style="text-align: right"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Var. Qty.2">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_ADD_VAR_QTY2" runat="server" Font-Size="11px" />
                                        <asp:HiddenField ID="ADD_VAR_QTY2" runat="server" />
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
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />
                    </td>
                </tr>
            </table>
        </div>

        <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
            ScrollBars="None">
            <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px; height: 80px">
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
