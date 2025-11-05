<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ROMain.aspx.vb" Inherits="INBOUND_RO_ROMain" %>
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

    <style type="text/css">
        .auto-style1 {
            font-family: "Arial" , "Helvetica" , "sans-serif";
            font-size: 14px;
            background-color: #d4d0c8;
            color: #193972;
            height: 25px;
        }
        .auto-style2 {
            height: 25px;
        }
    </style>

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

    function OpenDSCPRPT() {
        
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "DSCPReport", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=850,height=800,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "ro_code", document.getElementById("RO_CODE_HF").value);
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState("STORER_CODE")%>");
    }
    
    document.hiddenForm.action = "../../REPORT/DSCP_RPT/dscp_rpt_print.aspx";
    document.hiddenForm.target = "DSCPReport";
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
        setInterfaceDataToForm(document.hiddenForm, "EBS_PO_NO", document.getElementById("RO_EDI_PO_NO").value);
    document.hiddenForm.action = "RO_RECEIPT/stock_receipt.aspx";
    document.hiddenForm.target = "roPrint";
    document.hiddenForm.submit();
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
    function sumSubTotalPrice(ctrQTY,ctrPrice,Total)
    {
        var totalPrice, qty, price

        if (document.getElementById(ctrQTY).value != "") { qty = document.getElementById(ctrQTY).value; } else{ qty = 0; }
        if (document.getElementById(ctrPrice).value != "") { price = document.getElementById(ctrPrice).value; } else{ price = 0; }
        
        totalPrice = qty * price;

        document.getElementById(Total).value = totalPrice;
    }

    function OpenNotes() {
        
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "notes", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "DOC_TYPE", "IB_RO");
    setInterfaceDataToForm(document.hiddenForm, "DOC_NO", document.getElementById("RO_CODE_HF").value);
    setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
   
    document.hiddenForm.action = "../../ALERT/NotesMain.aspx";
    document.hiddenForm.target = "notes";
    document.hiddenForm.submit();
}
  </script>
    </div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
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
                            <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0" />
                        </td>
                        <td class="TITLE" align="right">
                            <asp:Button runat="server" ID="btnNOTE" text="Notes" CssClass="all_button" OnClientClick="javascript:OpenNotes();return false;" />
                            <input type="button" runat="server" id="btnAttach" value="Attachment" class="all_button" />
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
                                
                                <asp:Button ID="reOpenBtn" Text ="Re-Open" CssClass="all_button" runat ="server" Visible = "false" />
                                <input id="btnBack2" type="button" <%if Session("gLang") = "C" then %> value="返回" <% else %>value="Back" <%end if %> onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=IB_RO <%else %>../../IMPORT/UL_RO/UploadRO.aspx<%end if %>'"
                                    class="all_button" />
                                    
                            </td>
                            <td class="menuTD" align="right">
                                 <asp:Button ID="CancelBtn" Text ="Close" CssClass="all_button" runat ="server" />
                                 <asp:Button ID="btnPrintPalletLabel" Text="Print Pallet Label" CssClass="all_button" runat="server" />
                                <asp:Button ID="BtnGenBatchNo" Text="Generate Batch No." runat="server" CssClass="all_button"  />
                                <asp:Button ID="BtnCopy" Text="Copy WPO" CssClass="all_button" runat="server" OnClientClick="return confirm('A new PO wil be created and will be move forward to new PO page.\nAre you sure?');" />
                                <input type="button" value="Print Order" runat="server" class="all_button" id="btnPrintOrd" onclick="OpenOrd();" />
                                <input type="button" value="Print Item Label" runat="server" class="all_button" id="btnPrintLbl" onclick="OpenItemLbls();" />
                                <asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" />
                                <asp:Button ID="btnConfirm" Text="Confirm" CssClass="all_button" Visible="false" runat="server" />
                                <asp:button ID="btnDSCP" Text="Print Discrepancy Report" CssClass="all_button" runat="server" onclientclick="OpenDSCPRPT();return false;"  />
                                <asp:button ID="btnGenerate" Text="Generate Call-Off PO" CssClass="all_button" runat="server" OnClick="btnGenerate_Click" onClientClick="javascript:return confirm('Are you sure?');"  />
                       
                            </td>
                       </tr>
                    </table>                    
                </td>
            </tr>
            <tr>
                <td colspan="6" class="TITLE">
                    <asp:Table ID="linkBar" runat ="server" align="center" width="757"></asp:Table>                    
                </td>
                
            </tr>              
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_CODE" runat="server" >WPO Code</asp:Label>
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="RO_CODE" runat="server" />
                        <asp:hiddenField ID="RO_CODE_HF" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_STATUS" runat="server" >Status</asp:Label>
                    </font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="RO_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" >Storer</asp:Label>
                    </font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" MaxLength="20" AutoPostBack = true></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_EDI_PO_NO" runat="server" >EBS PO No.</asp:Label>
                    :</font>
                </td>
                <td>
                    <asp:TextBox ID="RO_EDI_PO_NO" runat="server" MaxLength="20" Width="150px" CssClass="REQUIRED"></asp:TextBox>
                    <asp:HiddenField runat="server" ID="RO_TRACK_NO" />
                </td>         
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_WH_CODE" runat="server" Text="Subinventory" />:</font>
                </td>
                <td>
                    <asp:dropdownlist ID="RO_WH_CODE" runat="server" CssClass="REQUIRED" />
                </td>         
            </tr>            
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_DATE" runat="server" >Date</asp:Label>
                    </font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="RO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                     <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />         
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_RCV_BY" runat="server" >Received By</asp:Label>
                    </font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="RO_RCV_BY" runat="server" MaxLength="20" Width="160px"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="auto-style1" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_PRJ_CODE" runat="server" >Project Code</asp:Label>
                    </font>
                </td>
                <td colspan="5" class="auto-style2">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                    <td width="50px" runat="server"><asp:DropDownList ID="PRJ_CODE" runat="server" MaxLength="20"></asp:DropDownList></td>
                    <!--
                    <td>&nbsp;<asp:Label ID="PRJ_NAME" runat="server" /></td>
                    -->
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_PO_TYPE" runat="server" >PO Type</asp:Label>
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:dropdownlist ID="PO_TYPE" runat="server" AutoPostBack="true">
                              <asp:ListItem Text="PO"  Value="PO" Selected />  
                              <asp:ListItem Text="NON-STOCK" Value="NS" />                                                                    
                        </asp:dropdownlist>                   
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_REF_NO" runat="server" >Contract No.</asp:Label>
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="RO_REF_NO" runat="server" MaxLength="30" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_BATCH_NO" runat="server" >Batch No.</asp:Label>
                    </font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="RO_BATCH_NO" runat="server" Width ="130px" MaxLength="20"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="auto-style1" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_PO_CAT" runat="server" >Category</asp:Label>
                    </font>
                </td>
                <td class="auto-style2">
                    <font size="2">
                        <asp:dropdownlist ID="PO_CAT" runat="server" AutoPostBack="true">
                              <asp:ListItem Text="FLOUR"  Value="FLOUR" Selected />  
                              <asp:ListItem Text="NON-STOCK" Value="NS" />                                                                    
                        </asp:dropdownlist>                   
                    </font>
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_ETD" runat="server" >ETD</asp:Label>
                    </font>
                </td>
                <td>
                    <asp:TextBox ID="RO_ETD" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_ETD" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_ETA" runat="server" >ETA</asp:Label>
                    </font>
                </td>
                <td>
                    <asp:TextBox ID="RO_ETA" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="RO_ETA" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_ISSUED_BY" runat="server" >Issued By</asp:Label>
                    </font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="RO_ISSUED_BY" runat="server" Width ="160px" MaxLength="50"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_STORER_TEL" runat="server" >Store&#39;s Tel.</asp:Label>
                    </font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="RO_STORER_TEL" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_STORER_FAX" runat="server" >Store&#39;s Fax</asp:Label>
                    </font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="RO_STORER_FAX" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_STORER_EMAIL" runat="server" >Store&#39;s Email</asp:Label>
                    </font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="RO_STORER_EMAIL" runat="server" MaxLength="30"></asp:TextBox>
                    </font>
                </td>
            </tr>        
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_STORER_CONT" runat="server" >Store&#39;s Contract</asp:Label>
                    </font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="RO_STORER_CONT" runat="server" MaxLength="30"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_RO_CTRY_ORIGIN" runat="server" Text="Country of Origin" /></font>:
                </td>
                <td nowrap colspan="3">
                    <font size="2">
                        <asp:DropDownList runat="server" ID="RO_CTRY_ORIGIN" Font-Size="12px"/><asp:TextBox ID="RO_DEST" runat="server" MaxLength="50" Visible="false"></asp:TextBox>
                    </font>
                </td>                
            </tr>   
            <tr>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label ID="lbl_RO_SEAL_NO" runat="server" text="Seal No." />:
                    </font>
                </td>
                <td>
                    <asp:TextBox runat="server" id="RO_SEAL_NO" MaxLength="100" />
                </td>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label ID="lbl_RO_CONTAINER_NO" runat="server" text="Container No." />:
                    </font>
                </td>
                <td>
                    <asp:TextBox runat="server" id="RO_CONTAINER_NO" MaxLength="50" />
                </td>
                <td class="LabelTD">
                    <font size="2">
                        <asp:Label ID="lbl_RO_CUST_INV_NO" runat="server" text="Customer Invoice No." />:
                    </font>
                </td>
                <td>
                    <asp:TextBox runat="server" id="RO_CUST_INV_NO" MaxLength="50" />
                </td>
            </tr>      
            <tr>
                 <td class="LabelTD">
                    <font size="2">
                        <asp:Label ID="lbl_RO_SHIP_MODE" runat="server" >Ship Mode</asp:Label>
                     :
                    </font>
                </td>
                <td colspan="5">
                    <asp:DropDownList runat="server" ID="RO_SHIP_MODE" />
                </td>
            </tr>       
            <tr>
                <td class="LabelTD" nowrap valign="top">
                    <font size="2">
                        <asp:Label ID="lbl_RO_REM" runat="server" >Remarks</asp:Label>
                    </font>
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
                    <%--<asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" visible="false" />--%>
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" />    
                    <asp:Button ID="btnCopyItem" runat="server" Text="Copy Items" CssClass="all_button" />
                </td>
            </tr>
            <tr>
                <td colspan="6">
                      <asp:UpdatePanel runat="server" ID="updtPnl_detailsList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="ro_code" HorizontalAlign="Left" ShowFooter="True">
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
                                    <asp:hiddenfield ID="rod_itm_code" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>                                                        
                            <asp:TemplateField HeaderText="Stock No.">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="vw_itm_sku_no" runat="server" Font-Size="11px" Width="150px"></asp:Label>
                                    <asp:textbox ID="itm_sku_no" runat="server" Font-Size="11px" Width="150px" Visible="false" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="150px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="vw_rod_itm_name" runat="server" Font-Size="11px" />
                                    <asp:textbox ID="rod_itm_name" runat="server" Visible="false" Font-Size="11px" CssClass="REQUIRED" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                          
                            <asp:TemplateField HeaderText="Pack Key">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="rod_pack_key" runat="server" Width="80px" Font-Size="10px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Distribution ID">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_pallet_no" runat="server" Width="70px" Font-Size="11px"></asp:TextBox>
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

                             <asp:TemplateField HeaderText="Transaction ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox ID="ROD_DOC_NO" runat="server" Font-Size="11px" Width="70px"/>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="PO Header ID">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox ID="ROD_SERIES_NO" runat="server" Font-Size="11px" Width="70px"/>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Carton No">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_carton_no" runat="server" Font-Size="11px" Width="70px"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Lot">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:ComboBox ID="rod_batch_no" runat="server" Font-Size="11px" RenderMode="block"
                                                  MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                    </asp:ComboBox>
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
                            <asp:TemplateField HeaderText="Lot No.">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:Label ID="ROD_MANU_DATE" runat="server" Font-Size="11px" Width="80"></asp:Label>
                                      <%--<asp:HiddenField ID="ROD_MANU_DATE" runat="server" />--%>
                                    <%-- <asp:TextBox ID="ROD_MANU_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                     <asp:ImageButton ID="btnCal02" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calMANU" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="ROD_MANU_DATE" PopupButtonID="btnCal02" Format="dd/MM/yyyy" />--%>
                                </ItemTemplate>
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
                            <asp:TemplateField HeaderText="Parent">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_itm_parent" runat="server" Font-Size="11px" Width="60px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <%-- <asp:TemplateField HeaderText="Series No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_series_no" runat="server" Font-Size="11px" Width="60px"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField> --%>            
                            <%--   <asp:TemplateField HeaderText="Location">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                     <asp:dropdownlist runat="server" ID="rod_location" Font-Size="11px"/>                                   
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>--%>
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
                            <%-- <asp:TemplateField HeaderText="Length">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_length" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="Width">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_width" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="Height">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_height" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>--%>
                            <%-- <asp:TemplateField HeaderText="Weight(KG)">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_kg" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="CBM">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_cbm" runat="server" Font-Size="11px" Width="50px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>--%>
                            <asp:TemplateField HeaderText="Status">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="rod_status" runat="server" Font-Size="11px" />
                                    <asp:HiddenField ID="rod_seq" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="rod_uom2" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty2">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="rod_qty2" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <%--<asp:TemplateField HeaderText="Currency">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="rod_curr" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="Rate">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="ROD_CURR_RATE" runat="server" Font-Size="11px" MaxLength="15" Width="60px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                            <%-- <asp:TemplateField HeaderText="Unit Price">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="ROD_UNIT_PRICE" runat="server" Font-Size="11px" MaxLength="15" Width="80px" style="text-align:right"  onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="Sub Total">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Textbox ID="PRICE_SUB_TOTAL" runat="server" Font-Size="11px" CssClass="READONLY" style="text-align:right" />
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                            <%--<asp:TemplateField HeaderText="HKD Equivalent">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:textbox ID="ROD_HKD_EQU" runat="server" Font-Size="11px" MaxLength="15" Width="80px" style="text-align:right" onkeypress="return maskKey(event);" />
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                            <%-- <asp:TemplateField HeaderText="On-Behalf">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:label ID="ROD_ON_BEHALF" runat="server" Font-Size="11px" style="text-align:left" />
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                            <asp:TemplateField HeaderText="Subinventory">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" ID="rod_wh_code" Font-Size="11px"  CssClass="REQUIRED" AutoPostBack="True" OnSelectedIndexChanged="rod_wh_code_SelectedIndexChanged"/>
                                </ItemTemplate>                          
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Location Code">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" ID="rod_location_code" Font-Size="11px"/>
                                </ItemTemplate>                          
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Insp. Req. Y/N">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="ROD_INSP_REQ" runat="server" />
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
                </ContentTemplate>
                          </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "C" then %> value="返回" <% else %>value="Back" <%end if %> onclick="Javascript:window.location='<% if ViewState("FrmUP") <>"Y" then %>../../cms_search.aspx?menu_code=IB_RO <%else %>../../IMPORT/UL_RO/UploadRO.aspx<%end if %>'"
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
