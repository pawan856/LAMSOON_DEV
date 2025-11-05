<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SCHKMain.aspx.vb" Inherits="OPERATION_STA_STAMain" MaintainScrollPositionOnPostback="true" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Stock Adjustment</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<style type="text/css">   
</style>
<div runat="server" id="DIVSCRIPT">
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

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1024,height=768,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_STKCHK_BAL");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        //setInterfaceDataToForm(document.hiddenForm, "mwh", WH);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2, seqList, ListNoList|CKD_CC_LIST_NO#txt_list_no, checkerList|CKD_CHECKER#txt_checker");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem() {
    document.myform.moduleAction.value = "SELECTIM";
    document.myform.submit();
}

function LocLookUp(lb_id, hd_id, wh_lb, wh_hd, csms_lb, csms_hd) {
    if(document.myform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "fun_code", 'OP_SCHK');
        setInterfaceDataToForm(document.hiddenForm, "mwh", document.myform.CK_WH.value);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_lb+ "|WHL, " + wh_hd + "|WHH, " + csms_lb + "|CSMSL, " + csms_hd + "|CSMSH");
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

    function ItemMastLookUp(STORER_CODE) {
        if (STORER_CODE == '')
        {
            alert('Please select the Storer first!');
            return false;
        }

        if(document.myform.editMode.value!="V")
        {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15");
            
            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM_STCHK");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItemMast()");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2, seqList, ListNoList|CKD_CC_LIST_NO#txt_list_no, checkerList|CKD_CHECKER#txt_checker");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "ItemLookUp";
            document.hiddenForm.submit();
        }
    }

    function selectedItemMast()
    {   
            document.myform.moduleAction.value = "SELECTIMMAST";
            document.myform.submit();
    }


    function maskKey(objEvent) {
        var iKeyCode;
        iKeyCode = objEvent.keyCode;
        if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
        return false;
    }

function calQty(rQty, loqid, lvqid, oqid, vqid) {
    var vQty;
    var oQty = eval("document.getElementById('" + oqid + "')").value;

    if (rQty == '') rQty = oQty;

    if (parseInt(oQty) > parseInt(rQty)) {
        vQty = (parseInt(oQty) - parseInt(rQty));
        eval("document.getElementById('" + lvqid + "')").innerHTML = "-" + vQty;
        eval("document.getElementById('" + vqid + "')").value = "-" + vQty;

        document.getElementById(lvqid).style.color = "red";

    } else if (parseInt(oQty) < parseInt(rQty)) {
        vQty = (parseInt(rQty) - parseInt(oQty));
        eval("document.getElementById('" + lvqid + "')").innerHTML = vQty;
        eval("document.getElementById('" + vqid + "')").value = vQty;

        document.getElementById(lvqid).style.color = "red";
    } else if (parseInt(oQty) == parseInt(rQty)) {
    eval("document.getElementById('" + lvqid + "')").innerHTML = 0;
    eval("document.getElementById('" + vqid + "')").value = 0;
    document.getElementById(lvqid).style.color = "black";
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
        //setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.CK_WH.value);

        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|1|L, " + hd_id + "|1, " + pallet_id + "|2|L, " + hpallet_id + "|2, " + lQty_id + "|3|L, " + qty_id + "|3, " + lbatch_no + "|4|L, " + batch_no + "|4, " + lexp_date + "|5|L, " + exp_date + "|5, " + lmanu_date + "|6|L, " + manu_date + "|6 ");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLocLookUp";
        document.hiddenForm.submit();
    }
}

function newChkList(flag) {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "CheckListMain" + flag, "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1100,height=768,left=5,top=15");


    setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);    
    
    setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
    setInterfaceDataToForm(document.hiddenForm, "ck_code", document.getElementById('CK_CODE').innerHTML);
    setInterfaceDataToForm(document.hiddenForm, "sheet_type", flag);
    document.hiddenForm.action = "CheckList.aspx";
    document.hiddenForm.target = "CheckListMain" + flag;
    document.hiddenForm.submit();
}

//function newChkListP() {
//    removeAllElementFromForm(document.hiddenForm);

//    window.open("", "CheckListMainP", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1100,height=768,left=5,top=15");


//    setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);    
//    
//    setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
//    setInterfaceDataToForm(document.hiddenForm, "ck_code", document.getElementById('CK_CODE').innerHTML);
//    setInterfaceDataToForm(document.hiddenForm, "sheet_type", "P");
//    document.hiddenForm.action = "CheckList.aspx";
//    document.hiddenForm.target = "CheckListMainP";
//    document.hiddenForm.submit();
//}

function newAdjForm() {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "AdjFormMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1250,height=850,left=5,top=15");
    
    setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);    
    
    setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
    setInterfaceDataToForm(document.hiddenForm, "ck_code", document.getElementById('CK_CODE').innerHTML);
    document.hiddenForm.action = "ADJ_FORM/adjForm.aspx";
    document.hiddenForm.target = "AdjFormMain";
    document.hiddenForm.submit();
}

    function reloadPage(alertMsg) {
        if (alertMsg)
            alert(alertMsg);

        removeAllElementFromForm(document.hiddenForm);
        setInterfaceDataToForm(document.hiddenForm, 'CK_CODE', document.getElementById("CK_CODE").innerHTML);

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("STORER_CODE").value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
        }
        
        document.hiddenForm.action = './SCHKMain.aspx';
        document.hiddenForm.target = '_self';
        document.hiddenForm.submit();
    }
</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <br />
    <script language="javascript">
        var xPos, yPos;
        var prm = Sys.WebForms.PageRequestManager.getInstance();

        function BeginRequestHandler(sender, args) {
            if ($get('<%=panelContainer.ClientID%>') != null) {
                // Get X and Y positions of scrollbar before the partial postback
                xPos = $get('<%=panelContainer.ClientID%>').scrollLeft;
                yPos = $get('<%=panelContainer.ClientID%>').scrollTop;
            }
        }

        function EndRequestHandler(sender, args) {
            if ($get('<%=panelContainer.ClientID%>') != null) {
                // Set X and Y positions back to the scrollbar
                // after partial postback
                $get('<%=panelContainer.ClientID%>').scrollLeft = xPos;
                $get('<%=panelContainer.ClientID%>').scrollTop = yPos;
            }
        }

        prm.add_beginRequest(BeginRequestHandler);
        prm.add_endRequest(EndRequestHandler);

    
    </script>
        <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="packKeyList" runat ="server" />    
    <asp:HiddenField ID="ListNoList" runat ="server" />   
    <asp:HiddenField ID="seqList" runat="server" /> 
    <asp:HiddenField ID="checkerList" runat="server" /> 
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
            <tr>
                <td colspan="4" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                    <asp:label ID="lheader" runat ="server" />
                     <a class="link" href="../SCHK/CC_sku_sample_list.xlsx" style="float:right;color:white;text-decoration-style:none!important;">Download</a>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />                    
                    </td>
                    <td class="menuTD" align="right">
                        <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />
                        <asp:FileUpload runat="server" ID="fu_SkuUpload" Width="200px" />
                                <asp:Button ID="btnImport" Text="Import" CssClass="all_button" runat="server" /> 
                    <asp:UpdatePanel runat="server" RenderMode="Inline" ID="udp1">
                            <ContentTemplate>
                                   
                                &nbsp;<asp:Button ID="btnRelease" runat="server" CssClass="all_button" Text="Release" />
                            </ContentTemplate>
                        </asp:UpdatePanel>                        
                    <asp:Button ID="btnPrint" Text="Print Cycle-Count List" runat="server" CssClass="all_button" OnClientClick="javascript:newChkList('N');return false;" />
                    <asp:Button ID="btnPrint3" Text="Print Cycle-Count List (Re-Check)" runat="server" CssClass="all_button" OnClientClick="javascript:newChkList('R');return false;" />
                    <asp:Button ID="btnPrint2" Text="Print Physical Cycle-Count List" runat="server" CssClass="all_button" OnClientClick="javascript:newChkList('P');return false;" />
                    
                    <%If Session("pagemode") <> "N" Then%>             
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" Visible="false" />
                        <asp:Button ID="btnCnt" Text="Counted" CssClass="all_button" runat="server" Visible="false" />
                        <asp:Button ID="btnRechk" Text="Re-Check" CssClass="all_button" runat="server"  Visible="false" OnClientClick="return confirm(&quot;Are you sure to start Re-Check?&quot;);"/>
                        
                        <asp:Button ID="btnGen" runat="server" Text="Generate Stock Adjust" CssClass="all_button" />
                        <asp:Button ID="btnPrintAdj" Text="Print Adjustment Form" runat="server" CssClass="all_button" OnClientClick="javascript:newAdjForm();return false;" />
                        <% End If%>

                        
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="CK_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_STATUS" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:hiddenfield ID="CK_STATUS" runat="server" />
                        <asp:Label ID="DSP_CK_STATUS" runat="server"  />
                        </font>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                         <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ck_wh" runat="server" /></font>
                </td>              
                <td runat="server">
                    <font size="2">
                    <asp:UpdatePanel runat="server" ID="UDPWH" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:DropDownList ID="CK_WH" runat="server" AutoPostBack="true" />
                        </ContentTemplate>
                    </asp:UpdatePanel>                         
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    &nbsp;
                </td>
                <td>
                   &nbsp;
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_SUB_WH" runat="server" /></font>
                </td>              
                <td id="Td5" runat="server">
                    <font size="2">
                         <asp:UpdatePanel runat="server" ID="UDPSWH" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:DropDownList ID="CK_SUB_WH" runat="server" >
                                <asp:ListItem Value="" Text="SELECT" />
                            </asp:DropDownList>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                    </font>
                </td>                            
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_TYPE" runat="server" /></font>
                </td>
                <td id="Td3" runat="server">
                    <font size="2">
                    <asp:UpdatePanel runat="server" ID="typeUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:dropdownlist ID="CK_TYPE" runat="server" CssClass="REQUIRED" AutoPostBack="true" />    
                        </ContentTemplate>
                    </asp:UpdatePanel>                    
                    </font>
                </td>
                <%-- <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ck_period" runat="server" text="Period" /></font>
                </td>
                <td id="Td4" runat="server">
                    <font size="2">
                    <asp:UpdatePanel runat="server" ID="periodUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:dropdownlist ID="ck_period" runat="server" AutoPostBack="true" >
                                <asp:ListItem Value="" Text="SELECT" />
                                <asp:ListItem Value="1YR" Text="1 Year" />
                                <asp:ListItem Value="2YR" Text="2 Year" />
                                <asp:ListItem Value="4YR" Text="4 Year" />
                            </asp:dropdownlist>
                        </ContentTemplate>
                    </asp:UpdatePanel>                    
                    </font>
                </td>--%>
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                <asp:UpdatePanel runat="server" ID="UpdatePanel5" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:TextBox ID="CK_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" AutoPostBack="true"></asp:TextBox>
                            <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                                ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                            <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                                TargetControlID="CK_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />  
                        </ContentTemplate> 
                </asp:UpdatePanel> 
                    
                </td>                
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ck_end_date" runat="server" Text="End Date" /></font>
                </td>
                <td>
                    <asp:UpdatePanel runat="server" ID="UpdatePanel2" RenderMode="Inline">
                       <ContentTemplate>
                          <asp:TextBox ID="ck_end_date" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                             <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                                  ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                          <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                                  TargetControlID="ck_end_date" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />  
                       </ContentTemplate>
                    </asp:UpdatePanel> 
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ck_start_date" runat="server" Text="Start Date" /></font>
                </td>
                <td width="150px">
                    <asp:UpdatePanel runat="server" ID="UpdatePanel1" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:TextBox ID="ck_start_date" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                            <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                                    ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                            <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                                    TargetControlID="ck_start_date" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />          
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_ACTUAL_DATE" runat="server" Text="Actual Date" /></font>
                </td>
                <td>
                
                            <asp:TextBox ID="CK_ACTUAL_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                               <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="../../images/calendar1.gif" 
                                    ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                            <asp:CalendarExtender ID="CalendarExtender5" runat="server" CssClass="ajax_calendar" 
                                    TargetControlID="CK_ACTUAL_DATE" PopupButtonID="ImageButton4" Format="dd/MM/yyyy" />  
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_COMPLT_DATE" runat="server" Text="Complete Date" /></font>
                </td>
                <td colspan="3">
                
                            <asp:TextBox ID="CK_COMPLT_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                               <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="../../images/calendar1.gif" 
                                    ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                            <asp:CalendarExtender ID="CalendarExtender4" runat="server" CssClass="ajax_calendar" 
                                    TargetControlID="CK_COMPLT_DATE" PopupButtonID="ImageButton3" Format="dd/MM/yyyy" />  
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_Total_Items" runat="server" Text="Total Items" /></font>
                </td>
                <td width="150px" colspan="3">
                    <asp:Label runat="server" id="total_items" /> 
                    <asp:UpdatePanel runat="server">
                        <ContentTemplate>
                            <asp:CheckBox runat="server" ID="CK_IS_Cable" Visible="false" />    
                        </ContentTemplate>
                    </asp:UpdatePanel>                    
                </td>
                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:dropdownlist runat="server" ID="CK_BY" />
                </td>                
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_RECHECK_BY" runat="server" Text="Re-Check By" /></font>
                </td>
                <td>
                    <asp:dropdownlist runat="server" ID="CK_RECHECK_BY" />
                </td>                
            </tr>
            <tr runat="server" id="lvTR" visible="false">
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ck_level" runat="server" Text="Level" /></font>
                </td>
                <td id="Td1" runat="server">
                    <font size="2">
                         <asp:DropDownList ID="CK_LEVEL" runat="server">
                            <asp:ListItem value="" text="Select" />
                            <asp:ListItem value="WH" text="Warehouse" />
                            <asp:ListItem value="FL" text="Floor" />
                            <asp:ListItem value="AR" text="Area" />
                         </asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_ck_item_level" runat="server" Text="Item Check Level" />
                    </font>
                </td>
                <td id="Td2" runat="server">
                    <font size="2">
                    <asp:DropDownList runat="server" ID="CK_ITEM_LEVEL">
                        <asp:ListItem Value="AL" text="All" />
                        <asp:ListItem Value="IT" Text="Item No" />
                    </asp:DropDownList>
                    </font>
                </td>
            </tr>
            <tr runat="server" id="refTR" visible="false" >
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_REF_NO" runat="server" /></font>
                </td>
                <td nowrap colspan="3">
                    <font size="2">
                        <asp:TextBox ID="CK_REF_NO" runat="server" MaxLength="30" Width="150px"></asp:TextBox>
                    </font>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap width="15%">
                    <font size="2">
                        <asp:Label ID="lbl_CK_REM" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="CK_REM" runat="server" rows="5" Width="523px" 
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
            </table>

         <asp:Button ID="btnExport" runat="server" text="Export to Excel" cssClass="all_button" />
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
                <td colspan="2" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:updatepanel runat="server" ID="LOOKUPDUP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:CheckBox runat="server" ID="chkITM" Text="Select from Item master" Visible="false" />
                            <asp:Button ID="selectItemBtn" runat="server" text="Select from Item Balance" cssClass="all_button" />                            
                            <asp:Button ID="selectItemMastbtn" runat="server" text="Select from Item Master" cssClass="all_button" />
                            <asp:Button ID="addBlank" runat="server" text="Add Blank Item"   cssClass="all_button" />
                            <asp:Button ID="updateBOOK" runat="server" Text="Update Recheck Booked no." cssClass="all_button" Visible = "false" />
                            <div style="display:inline; float:right;">
                                <asp:Button runat="server" ID="btnDelAll" Text="Delete All Item" CssClass="all_button" />
                            </div>
                        </ContentTemplate>
                    </asp:updatepanel>
                </td>
                <td colspan="2" class="menuTD">
                    <asp:Label runat="server" id="lbl_filter" Text="Filter By Status:" />
                    <asp:updatepanel runat="server" ID="Updatepanel4" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:dropdownlist runat="server" ID="status_filter" AutoPostBack="true">
                                <asp:ListItem Value="" Text="SELECT" />
                                <asp:ListItem Value="FC" Text="First Check" />
                                <asp:ListItem Value="RE" Text="Re-Check" />
                                <asp:ListItem Value="CD" Text="Counted" />
                            </asp:dropdownlist>
                        </ContentTemplate> 
                    </asp:updatepanel> 
                </td>
            </tr>
            <tr>
                <td colspan="4">
                 <asp:Button ID="Button1" runat="server" Text="Add" CssClass="all_button" Visible = "false" />
                    <asp:updatepanel runat="server" ID="Updatepanel3" RenderMode="Inline">
                    <ContentTemplate>
                    <asp:panel runat="server" id="panelContainer" style="max-Height:720px" ScrollBars="Auto">
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="ckD_seq">
                           <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnCount" name="btnCount" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="COUNT" Text="CNT" CssClass="all_button" Font-Bold="false" style="padding:0;" />                                                                        
                                </ItemTemplate>
                                <ControlStyle Width="40px" />
                                <ControlStyle Width="40px" />
                                <HeaderStyle Width="45px" />
                            </asp:TemplateField>
                             <asp:BoundField DataField="CKD_seq" HeaderText="Seq No." >
                                <ControlStyle Width="30px"  />
                                <ControlStyle Width="30px"  />
                                <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Item code">
                                <ControlStyle Width="100px" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" runat="server" Font-Size="11px" Width="100px"></asp:Label>
                                    <asp:HiddenField runat="server" ID="CKD_SEQ" />
                                    <asp:hiddenfield ID="ckd_itm_code" runat="server" />
                                    <asp:HiddenField ID="ckd_pack_key" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Subinventory">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:label ID="dsp_WH_CODE" runat="server" Font-Size="11px" Width="30px"></asp:label>
                                    <asp:HiddenField runat="server" ID="WH_CODE" />
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="EDI Code">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:label ID="dsp_BN_CSMS_CODE" runat="server" Font-Size="11px"></asp:label>
                                    <asp:HiddenField runat="server" ID="BN_CSMS_CODE" />
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ItemTemplate>
                                    <asp:Label ID="itm_name" runat="server" Font-Size="11px" Width="120px"></asp:Label>
                                    <asp:textbox ID="itm_name_textbox" runat="server"  Width="98%" Visible="false" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Serial No.">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="dsp_CKD_SERIAL_NO" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" id="CKD_SERIAL_NO" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Lot">
                                <HeaderStyle />
                                <ItemTemplate>
                                    <asp:label ID="disp_ckd_batch_no" runat="server" Font-Size="11px" />
                                    <asp:hiddenfield ID="ckd_batch_no" runat="server" />
                                    <asp:HiddenField ID="ckd_pallet_no" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:BoundField runat="server" DataField="ITM_UOM" HeaderText="UOM" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                            <asp:TemplateField HeaderText="Org.<br>Qty">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_ckd_org_qty" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                    <asp:HiddenField ID="ckd_org_qty" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Checked<br>Qty">
                                <ItemTemplate>
                                    <asp:TextBox ID="ckd_rev_qty" runat="server" Font-Size="11px" Width="60px" MaxLength="15" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ReChecked<br>Booked">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_CKD_BOOK_QTY" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                    <asp:HiddenField ID="CKD_BOOK_QTY" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ReChecked<br>Actual">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_ACTUAL_QTY" runat="server" Font-Size="11px" Width="60px" MaxLength="15" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Variance<br>Qty">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_ckd_var_qty" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                    <asp:HiddenField ID="ckd_var_qty" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Drum ID" Visible="false">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="dsp_CKD_DRUM_ID" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" id="CKD_DRUM_ID" />
                              </ItemTemplate>                              
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Org. Qty2" Visible="false">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="dsp_CKD_ORG_QTY2" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" id="CKD_ORG_QTY2" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM2" Visible="false">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="dsp_CKD_UOM2" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" id="CKD_UOM2" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Chk. Qty2" Visible="false">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_REV_QTY2" runat="server" Width="60px" Font-Size="10px" MaxLength="15"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ReChk. Booked2">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_CKD_BOOK_QTY2" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                    <asp:HiddenField ID="CKD_BOOK_QTY2" runat="server" />                                    
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ReChk. Actual2">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_ACTUAL_QTY2" runat="server" Width="60px" Font-Size="10px" MaxLength="15"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Var. Qty2" Visible="false">
                              <ItemTemplate>
                                  <asp:Label ID="dsp_ckd_var_qty2" runat="server" Font-Size="11px" Width="60px" style="text-align:right"></asp:Label>
                                  <asp:HiddenField ID="ckd_var_qty2" runat="server" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>

                             <asp:TemplateField HeaderText="Type" Visible="false">
                                <ItemTemplate>
                                    <asp:DropDownList runat="server" ID="CKD_TYPE">
                                        <asp:ListItem Value="" Text="SELECT" />
                                        <asp:ListItem Text="Pre-Allocated" Value="A" />
                                        <asp:ListItem Text="UnKnown" Value="U" />
                                    </asp:DropDownList>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Remarks">
                                <ItemTemplate>
                                    <asp:TextBox ID="ckd_rem" runat="server" Font-Size="11px" MaxLength="200"></asp:TextBox>
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Status">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="DSP_CKD_STATUS" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" ID="CKD_STATUS" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Checker" Visible="false">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_CHECKER" runat="server" Width="80px" Font-Size="11px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>                                
                            </asp:TemplateField>
                           <%-- <asp:BoundField runat="server" DataField="CKD_BY" HeaderText="Chk By" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                            <asp:BoundField runat="server" DataField="CKD_CC1_DATE" HeaderText="Chk Date" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                            <asp:BoundField runat="server" DataField="CKD_RECHECK_BY" HeaderText="ReChk By" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />
                            <asp:BoundField runat="server" DataField="CKD_CC_DATE" HeaderText="ReChk Date" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" />--%>


                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" />
                                    <asp:HiddenField runat="server" ID="mFlag" />
                                </ItemTemplate>
                                <ControlStyle Width="50px"></ControlStyle>
                                <HeaderStyle Width="50px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Loc.">
                                <ItemTemplate>
                                   <asp:Label ID="dsp_ckd_loc" width="100px" runat="server" Font-Size="11px" />
                                   <asp:HiddenField ID="ckd_loc" runat="server" />
                                   <asp:HiddenField ID="from_imast" runat="server" />
                                   <asp:Imagebutton ID="Image_Loc_LookUp" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand" align="absmiddle" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" Wrap="false" />
                            </asp:TemplateField>                 
                            <asp:TemplateField HeaderText="Adj. No">
                                 <ItemTemplate>
                                    <asp:HyperLink runat="server" ID="CKD_AD_CODE" Font-Size="11px" Target="_self" />
                                 </ItemTemplate>                                 
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Adj. Remarks">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_ADJ_REM" runat="server" Font-Size="11px" MaxLength="500"></asp:TextBox>
                                </ItemTemplate>                                
                            </asp:TemplateField>



                            <asp:TemplateField HeaderText="List No.">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_CC_LIST_NO" runat="server" Font-Size="11px" Width="30px" style="text-align:right"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ItemTemplate>
                                    <asp:Label ID="dsp_ckd_pack_key" runat="server" Font-Size="11px" Width="30px" ></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                           
                            <asp:TemplateField HeaderText="Count Date" Visible="false" >
                                <ItemStyle wrap="false" />
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_CC_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                    <asp:ImageButton ID="btnCal" runat="server" ImageUrl="../../images/calendar1.gif" 
                                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                                            TargetControlID="CKD_CC_DATE" PopupButtonID="btnCal" Format="dd/MM/yyyy" />          
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Full Drum" Visible="false">
                              <ItemTemplate>
                                  <asp:dropdownlist runat="server" id="CKD_FULL_DRUM" Font-Size="11px">
                                  <asp:ListItem Text="Y" Value="Y" />
                                  <asp:ListItem Text="N" Value="N" Selected />
                                  </asp:dropdownlist>
                              </ItemTemplate>                              
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Level" Visible="false">
                              <ItemTemplate>
                                  <asp:Label runat="server" id="dsp_CKD_DRUM_LEVEL" Font-Size="11px" />
                                  <asp:hiddenfield runat="server" id="CKD_DRUM_LEVEL" />
                              </ItemTemplate>
                              <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Witness By" Visible="false">
                                <ItemTemplate>
                                    <asp:TextBox ID="CKD_WITNESS" runat="server" Width="80px" Font-Size="11px" MaxLength="50"></asp:TextBox>
                                </ItemTemplate>                                
                            </asp:TemplateField>
                            <asp:BoundField runat="server" DataField="CKD_IN_PDA" HeaderText="Chked in PDA" HeaderStyle-HorizontalAlign="Left" ItemStyle-Font-Size="11px" Visible="false" />

                            </Columns>
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                    </asp:GridView>
                    </asp:panel>
                    </ContentTemplate>
                    </asp:updatepanel>                    
                </td>
            </tr>
            <tr>
                <td colspan="4" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
            </tr>
        </table>
        <asp:HiddenField runat="server" ID="dummy" />

    <asp:Panel ID="PnlSC" runat="server" CssClass="modalPopup" style="display: none" ><!---->
        <table width="840px">
           <tr>
                <td class="TITLE">
                    <asp:Panel runat="Server" ID="PanelDrag1" Style="cursor: move;">
                    <font size="2">Generate Stock Adjust</font>
                    </asp:Panel>
                </td>
           </tr>
           <tr>
                <td>
                   <font size="2">
                    Stock Adjustment record will be generated for the following item:
                   </font>
                </td>
           </tr>
           <tr>
              <td>
                   <asp:GridView ID="GVAdj" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                           <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                        <Columns>
                            <asp:BoundField DataField="itm_sku_no" HeaderText="Stock No." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="itm_name" HeaderText="Item Name" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_batch_no" HeaderText="Batch No." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_loc" HeaderText="Loc." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_org_qty" HeaderText="Org Qty." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_rev_qty" HeaderText="Verified Qty." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_var_qty" HeaderText="Variance Qty." >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_org_qty2" HeaderText="Org Qty.2" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_rev_qty2" HeaderText="Verified Qty.2" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ckd_var_qty2" HeaderText="Variance Qty.2" >
                                <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                   </asp:GridView> 
              </td>
           </tr>
           <tr>
                <td class="TITLE" align="left">
                    <asp:hiddenfield runat="server" ID="CAction" />
                    <asp:Button runat="server" ID="btnPnlSave" Text="Confirm" CssClass="all_button" />
                    <asp:Button runat="server" ID="btnPnlClose" Text= "Cancel" CssClass="all_button" />
                </td>
           </tr>
        </table>
    </asp:Panel> 
    <asp:ModalPopupExtender ID="pnlSC_ModalPopupExtender" runat="server"
        DynamicServicePath="" 
        Enabled="True" 
        TargetControlID="dummy" 
        PopupControlID="pnlSC"
        BackgroundCssClass="modalBackground"
        DropShadow="true"         
        CancelControlID="btnPnlClose"
        Y="50"
        PopupDragHandleControlID="PanelDrag1" RepositionMode="None">
    </asp:ModalPopupExtender>

    </div>
     <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>     
     <asp:UpdatePanel runat="server" ID="BinListUDP" RenderMode="Inline">
        <ContentTemplate>
        <asp:HiddenField runat="server" ID="dummy2" />
        <asp:Panel ID="OtherItmList" runat="server" CssClass="modalPopup" style="display: none"><!-- -->
        <div id="Div2" runat="server" style="max-height: 7200px; overflow: auto;">
        <table width="840px">
           <tr>
                <td class="TITLE">
                    <asp:Panel runat="Server" ID="DragPan2" Style="cursor: move;">
                    <font size="2">List Of bin Location</font>
                    </asp:Panel>
                    <div style="float:right">
                        <asp:Label runat="server" Font-Size="11px" Text="Total no. of bin location item:"  />
                        <asp:Label runat="server" Font-Size="11px" ID="TOTAL_NO_OF_ITM" />
                    </div>
                </td>
           </tr>
           <tr>
                <td>
             <%--<asp:panel runat="server" id="panel2" style="max-Height:200px" ScrollBars="Auto">--%>
                  <asp:GridView ID="GVBin" runat="server" Height="10px" Width="100%" Font-Names="Arial"  AllowPaging="True" PageSize="20"
                     Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                     BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                     CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                        <RowStyle CssClass="GV" />
                        <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                     <Columns>
                         <asp:TemplateField ItemStyle-HorizontalAlign="Right" Visible="false" >
                           <ItemTemplate>
                                <asp:Label runat="server" ID="RowIndex" Font-size="11px" />
                           </ItemTemplate>
                         </asp:TemplateField>
                         <asp:BoundField DataField="ILOC_LOC" HeaderText="Location" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                         </asp:BoundField>
                         <asp:BoundField DataField="BN_CSMS_CODE" HeaderText="EDI Code" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                         </asp:BoundField>
                         <asp:BoundField DataField="NO_OF_COUNT" HeaderText="No of Item" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                         </asp:BoundField>
                     </Columns>
                      <PagerTemplate>
                      <asp:UpdatePanel runat="server" ID="pagerUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <table cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td align="center">
                                        <asp:ImageButton ID="pager_first" ImageUrl="../../images/arrow_frist.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="First" />
                                        <asp:ImageButton ID="pager_previous" ImageUrl="../../images/arrow_previous.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Prev" />
                                        <asp:Label ID="lblPager" Text="Page : " Font-Bold="true" style="vertical-align:middle" runat="server" />
                                        <asp:DropDownList ID="pager_select" AutoPostBack="true" Width="50" style="vertical-align:middle" runat="server" OnSelectedIndexChanged="GVBin_PageIndexChanged" />
                                        <asp:ImageButton ID="pager_next" ImageUrl="../../images/arrow_next.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Next" />
                                        <asp:ImageButton ID="pager_last" ImageUrl="../../images/arrow_last.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Last" />                   
                                    </td>
                                </tr>
                            </table>
                        </ContentTemplate>
                      </asp:UpdatePanel>                        
                      </PagerTemplate>
                  </asp:GridView>
               <%--   </asp:panel>--%>
                </td>
           </tr>
           <tr>
            <td>&nbsp;</td>
           </tr>
           <tr>
                <td class="TITLE">
                    <asp:Panel runat="Server" ID="Panel1" Style="cursor: move;">
                    <font size="2">List Of Other bin Location Item</font>
                    </asp:Panel>
                    <div style="float:right">
                        <asp:Label ID="lbl_Other" runat="server" Font-Size="11px" Text="Total no. of other location item:"  />
                        <asp:label runat="server" Font-Size="11px" ID="TOTAL_NO_OF_OTHER" />
                    </div>
                </td>
           </tr>
           <tr>
                <td>
                <%--<asp:panel runat="server" id="panel3" style="max-Height:200px" ScrollBars="Auto">--%>
                   <asp:GridView ID="GVOther" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                     Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                     BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" AllowPaging="true" PageSize="20"
                     CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                        <RowStyle CssClass="GV" />
                        <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                     <Columns>
                         <asp:TemplateField ItemStyle-HorizontalAlign="Right" Visible="false">
                           <ItemTemplate>
                                <asp:Label runat="server" ID="RowIndex" Font-size="11px" />
                           </ItemTemplate>
                         </asp:TemplateField>
                         <asp:BoundField DataField="ILOC_LOC" HeaderText="Location" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                         </asp:BoundField>
                         <asp:BoundField DataField="BN_CSMS_CODE" HeaderText="CSMS Code" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Left" />
                         </asp:BoundField>
                         <asp:BoundField DataField="NO_OF_COUNT" HeaderText="No of Item" >
                             <ItemStyle Font-Size="11px" HorizontalAlign="Right" />
                         </asp:BoundField>
                     </Columns>
                     <PagerTemplate>
                      <asp:UpdatePanel runat="server" ID="pagerUDP2" RenderMode="Inline">
                        <ContentTemplate>
                            <table cellspacing="0" cellpadding="0" border="0" width="100%">
                                <tr>
                                    <td align="center">
                                        <asp:ImageButton ID="pager_firstO" ImageUrl="../../images/arrow_frist.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="First" />
                                        <asp:ImageButton ID="pager_previousO" ImageUrl="../../images/arrow_previous.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Prev" />
                                        <asp:Label ID="lblPagerO" Text="Page : " Font-Bold="true" style="vertical-align:middle" runat="server" />
                                        <asp:DropDownList ID="pager_selectO" AutoPostBack="true" Width="50" style="vertical-align:middle" runat="server" OnSelectedIndexChanged="GVOther_PageIndexChanged" />
                                        <asp:ImageButton ID="pager_nextO" ImageUrl="../../images/arrow_next.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Next" />
                                        <asp:ImageButton ID="pager_lastO" ImageUrl="../../images/arrow_last.png" Height="24" Width="24" style="vertical-align:middle" runat="server" CommandName="Page" CommandArgument="Last" />                   
                                    </td>
                                </tr>
                            </table>
                        </ContentTemplate>
                      </asp:UpdatePanel>
                      </PagerTemplate>
                  </asp:GridView>
               <%--   </asp:panel>--%>
                </td>
           </tr>
           <tr>
               <td>
                     <asp:Label runat="server" id="sql_debug" Font-Size="11px" />
               </td>           
           </tr>
           <tr>
               <td>
                    <asp:button runat="server" ID="btnPnlClose2" Text="Close" CssClass="all_button" /> 
               </td>           
           </tr>
           </table>
           </div>
         </asp:Panel> 

             <asp:GridView ID="GridTableData" runat="server"  Visible="false"></asp:GridView>
 
         <asp:ModalPopupExtender ID="OtherMDExt" runat="server"
            DynamicServicePath="" 
            Enabled="True" 
            TargetControlID="dummy2" 
            PopupControlID="OtherItmList"
            BackgroundCssClass="modalBackground"
            DropShadow="true"         
            CancelControlID="btnPnlClose2"
            Y="50"
            PopupDragHandleControlID="DragPan2" RepositionMode="None">
        </asp:ModalPopupExtender>
        </ContentTemplate>        
     </asp:UpdatePanel>     
    <asp:HiddenField ID="IMP_CODE" runat="server" />
    <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
