<%@ Page Language="VB" AutoEventWireup="false" CodeFile="STFMain.aspx.vb" Inherits="OPERATION_STF_STFMain" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Stock Transfer</title>
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

        function ItemLookUp(STORER_CODE, WH_CODE, MWH_CODE) {
            if (STORER_CODE == '') {
                alert('Please select the Storer first!');
                return false;
            }

            if (WH_CODE == '' && MWH_CODE == '') {
                alert('Please select the From Warehouse!');
                return false;
            }

            if (document.myform.editMode.value != "V") {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=600,left=5,top=15");

                setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_ILOC_BAL");
                setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
                setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
                setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
                setInterfaceDataToForm(document.hiddenForm, "wh", WH_CODE);
                setInterfaceDataToForm(document.hiddenForm, "mwh", MWH_CODE);
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

        function selectedTQ() {
            document.myform.moduleAction.value = "SELECTTQ";
            document.myform.submit();
        }

        function selectedToWH(rowIdx) {
            document.myform.selectedrowIndex.value = rowIdx;
            document.myform.moduleAction.value = "SELECTTOWH";
            document.myform.submit();
        }

        function LocLookUp(seq, rowIdx, lb_id, hd_id, mWH) {
            //if (document.myform.editMode.value != "V") {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "locLookUp" + seq, "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=700,left=5,top=15,resizable=yes");

            if (seq == 1) {
                setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TR_WH_FR.value);
            } else {
                setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.TR_WH_TO.value);

            }

            //setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedToWH('" + rowIdx + "')");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

            if (mWH == 'F') {
                setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.TR_WH_FR.value);
            }

            if (mWH == 'T') {
                setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.TR_WH_TO.value);
            }

            setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
            document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
            document.hiddenForm.target = "locLookUp" + seq;
            document.hiddenForm.submit();
            //}
        }

        function ItemLocLookUp(STORER_CODE, ITEM_CODE, PACK_KEY, lb_id, hd_id, pallet_id, hpallet_id, qty_id, batch_no,
            exp_date, manu_date, QTY2, DRUM_ID, DRUM_LEVEL, ILBS_SERIAL_NO) {
            if (STORER_CODE == '') {
                alert('Please select the Storer first!');
                return false;
            }

            if (document.myform.editMode.value != "V") {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "ItemLocLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,left=5,top=15");



                setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_LOC");
                setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");

                setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
                setInterfaceDataToForm(document.hiddenForm, "ic", ITEM_CODE);
                setInterfaceDataToForm(document.hiddenForm, "pKey", PACK_KEY);
                setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.TR_WH_FR.value);

                setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + qty_id + "|3, " + batch_no + "|4, " + exp_date + "|5, " + manu_date + "|6, "
                    + QTY2 + "|7, " + DRUM_ID + "|8, " + DRUM_LEVEL + "|9, " + ILBS_SERIAL_NO + "|10");
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

        function TQLookUp() {
            if (document.getElementById("STORER_CODE").value == '') {
                alert('Please select the Storer first!');
                return false;
            }


            if (document.myform.editMode.value != "V") {

                removeAllElementFromForm(document.hiddenForm);

                window.open("", "TQLookup", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=750,left=5,top=15");

                setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_TQ");
                setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
                setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedTQ()");
                setInterfaceDataToForm(document.hiddenForm, "sc", document.getElementById("STORER_CODE").value);
                setInterfaceDataToForm(document.hiddenForm, "pItemList", "TQList");
                document.hiddenForm.action = "../../cms_search.aspx";
                document.hiddenForm.target = "TQLookup";
                document.hiddenForm.submit();
            }
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
        <asp:HiddenField ID="TQList" runat="server" />
        <asp:HiddenField ID="TQSeqList" runat="server" />
        <asp:HiddenField ID="selectedrowIndex" runat="server" />
        <div id="div1">
            <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
                <tr>
                    <td colspan="3" class="TITLE" align="left">
                        <asp:Label ID="lheader" runat="server" Text="Stock Trasnfer Maintenance" />
                    </td>
                    <td colspan="1" class="TITLE" align="right">
                        <asp:Table ID="cmBar" runat="server" border="0" CellSpacing="0" CellPadding="0"></asp:Table>
                        <input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" />
                        <asp:Button ID="btnsearch" Text="Search" CssClass="all_button" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="menuTD">
                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                            <tr>
                                <td class="menuTD" align="left">
                                    <asp:Button ID="selectSTQ" runat="server" Text="Select IST Request" CssClass="all_button" Visible="false" />
                                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                                   
                                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                                        class="all_button" />
                                </td>
                                <td class="menuTD" align="right">
                                     <asp:Button ID="CancelBtn" Text="Cancel" CssClass="all_button" runat="server" />
                                    <%If Viewstate("pagemode") <> "N" Then%>
                                    <asp:Button ID="btnGenWO" CssClass="all_button" Text="Generate WO" runat="server" Visible="false" />
                                    <asp:Button ID="btnIssue" Text="Issue" CssClass="all_button" runat="server" />
                                    <asp:Button ID="btnRcv" Text="Receive" CssClass="all_button" runat="server" />
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
                            <asp:Label ID="lbl_TR_CODE" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:Label ID="TR_CODE" runat="server" />
                        </font>
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_STATUS" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:Label ID="TR_STATUS" runat="server" /></font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack="true"></asp:DropDownList>
                        </font>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_TR_EDI_RFT_NO" runat="server" /></font>
                    </td>
                    <td id="Td1" colspan="3" runat="server">
                        <font size="2">
                            <asp:TextBox ID="TR_EDI_RFT_NO" runat="server" MaxLength="10"></asp:TextBox>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_DATE" runat="server" /></font>
                    </td>
                    <td width="150px">
                        <asp:TextBox ID="TR_DATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                        <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif"
                            ImageAlign="Middle" BorderWidth="0" Style="padding-left: 4px" />
                        <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar"
                            TargetControlID="TR_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_BY" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="TR_BY" runat="server" MaxLength="20"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_BATCH_NO" runat="server" /></font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:TextBox ID="TR_BATCH_NO" MaxLength="20" runat="server" /></font>
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_REF_NO" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <font size="2">
                            <asp:TextBox ID="TR_REF_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_WH_FR" runat="server" /></font>
                    </td>
                    <td width="26%">
                        <font size="2">
                            <asp:DropDownList ID="TR_WH_FR" runat="server" AutoPostBack="true" />
                        </font>
                    </td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_WH_TO" runat="server" /></font>
                    </td>
                    <td width="26%">
                        <font size="2">
                            <asp:DropDownList ID="TR_WH_TO" runat="server" AutoPostBack="true"></asp:DropDownList>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%"></td>
                    <td id="Td2" width="26%" runat="server"></td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_TO_LOC" runat="server" /></font>
                    </td>
                    <td id="Td3" width="26%" runat="server">
                        <font size="2">
                            <asp:Label ID="dsp_TR_TO_LOC" runat="server" /></font>
                        <asp:HiddenField ID="TR_TO_LOC" runat="server" />
                        <asp:Image ID="Image_Loc_LookUp_TR_TO_LOC" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" Style="border-width: 0px; cursor: hand" align="absmiddle" />
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%"></td>
                    <td id="Td4" width="26%" runat="server"></td>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_TO_DRUM" runat="server" /></font>
                    </td>
                    <td id="Td5" width="26%" runat="server">
                        <font size="2">
                            <asp:TextBox ID="TR_TO_DRUM" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                        </font>
                    </td>
                </tr>
                <tr runat="server" id="TQ_TR">
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_TQ_NO" runat="server" Text="RFT Code" />:</font>
                    </td>
                    <td colspan="3">
                        <font size="2">
                            <asp:Label runat="server" ID="TR_TQ_NO" />
                        </font>
                    </td>
                </tr>
                <tr runat="server" id="WO_TR" visible="false">
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="Label1" runat="server" Text="RFID WO Code" />:</font>
                    </td>
                    <td>
                        <font size="2">
                            <asp:LinkButton ruant="server" ID="WO_TR_CODE" runat="server" />
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap width="15%">
                        <font size="2">
                            <asp:Label ID="lbl_TR_REM" runat="server" /></font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="TR_REM" runat="server" Height="117px" Width="523px"
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
                        <asp:UpdatePanel runat="server" ID="LOOKUPUDP" RenderMode="Inline">
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
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="trd_seq"  HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                 
                        <Columns>
                            <asp:TemplateField HeaderText="No.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>                                   
                                    <asp:TextBox ID="trd_seq" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" BorderWidth="0" BackColor="Transparent" readonly="true" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <%-- <asp:BoundField DataField="trd_seq" HeaderText="No." >
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" />
                            </asp:BoundField>--%>
                               <asp:TemplateField HeaderText="Item Code">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <%--<asp:TextBox ID="itm_name" runat="server" Font-Size="11px" Width="110px" MaxLength="100"></asp:TextBox>--%>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="itm_sku_no" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                                    <asp:HiddenField runat="server" id="itm_code" />
                                    <asp:HiddenField runat="server" id="TRD_EXPIRY_DATE" />
                                    <asp:HiddenField runat="server" id="TRD_MANU_DATE" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="pack_key" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" width="60px" id="TRD_ITM_NAME" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>    
                            <asp:TemplateField HeaderText="Lot">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" id="trd_batch_no_fr" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Pallet No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                        <table border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_trd_pallet_no_fr" runat="server" Font-Size="11px" />
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="trd_pallet_no_fr" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="TRD_UOM" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Req. Qty" Visible="false">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Font-size="11px" id="trd_req_qty" style="text-align:right" CssClass="READONLY" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Org Qty">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Font-size="11px" id="trd_org_qty" style="text-align:right" CssClass="READONLY" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="trd_qty" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Serial No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Width="50px" Font-size="11px" id="TRD_SERIAL" MaxLength="80" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label runat="server" Font-size="11px" id="TRD_UOM2" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Org Qty2">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Font-size="11px" id="trd_org_qty2" style="text-align:right" CssClass="READONLY" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty2">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="trd_qty2" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="From Drum ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Font-size="11px" id="TRD_DRUM_ID_FR" Width="50px" CssClass="READONLY" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="From Drum Lv.">
                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Font-size="11px" id="TRD_DRUM_LV_FR" Width="50px" CssClass="READONLY"  />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Width="40px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="From Loc.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_trd_loc_fr" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td>
                                                    <asp:Image ID="Image_Loc_LookUp_FR" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="trd_loc_fr" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="To Drum ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Width="50px" Font-size="11px" id="TRD_DRUM_ID_TO" MaxLength="80" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="To Drum Lv.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox runat="server" Width="50px" Font-size="11px" id="TRD_DRUM_LV_TO" MaxLength="10" onkeypress="return maskKey(event);" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="To Loc.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                            <tr>
                                                <td>
                                                     <asp:Label ID="dsp_trd_loc_to" width="100px" runat="server" Font-Size="11px" />
                                                </td>
                                                <td>
                                                    <asp:Image ID="Image_Loc_LookUp_To" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />     
                                                </td>
                                            </tr>
                                        </table>                  
                                        <asp:HiddenField ID="trd_loc_to" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="110px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="trd_rem" runat="server" Width="110px" Font-Size="10px" MaxLength="200"></asp:TextBox>
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
