<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SOMain.aspx.vb" Inherits="OUTBOUND_SO_SOMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
 <title>Shipment Order</title>
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

        function ItemLookUp(STORER_CODE) {
            if (STORER_CODE == '') {
                alert('Please select the Storer first!');
                return false;
            }

            if (document.myform.editMode.value != "V") {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=700,height=500,left=5,top=15");

                setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM");
                setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
                setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
                setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
                setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1");
                document.hiddenForm.action = "../../cms_search.aspx";
                document.hiddenForm.target = "ItemLookUp";
                document.hiddenForm.submit();
            }
        }

        function DOLookUp(STORER_CODE) {
            if (STORER_CODE == '') {
                alert('Please select the Storer first!');
                return false;
            }

            if (document.myform.editMode.value != "V") {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "DOLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");

                setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_DO");
                setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
                setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedDO()");
                setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
                setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, seqList|2");
                document.hiddenForm.action = "../../cms_search.aspx";
                document.hiddenForm.target = "DOLookUp";
                document.hiddenForm.submit();
            }
        }

        function selectedDO() {
            document.myform.moduleAction.value = "SELECTDO";
            document.myform.submit();
        }

        function checkSB(STORER_CODE) {
            if (document.myform.editMode.value != "V") {
                removeAllElementFromForm(document.hiddenForm);

                window.open("", "sbLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");

                setInterfaceDataToForm(document.hiddenForm, "menu_code", "INQ_001");
                setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
                setInterfaceDataToForm(document.hiddenForm, "ctemp", true);
                setInterfaceDataToForm(document.hiddenForm, "screadonly", true); 

                document.hiddenForm.action = "../../cms_search.aspx";
                document.hiddenForm.target = "sbLookUp";
                document.hiddenForm.submit();
            }
        }

        function chgShipMode(mode) {
            if (mode == "SEA") {
                document.getElementById("seaTR").style.display = "";
                document.getElementById("airTR1").style.display = "none";
                document.getElementById("airTR2").style.display = "none";
            } else if (mode == "AIR") {
                document.getElementById("seaTR").style.display = "none";
                document.getElementById("airTR1").style.display = "";
                document.getElementById("airTR2").style.display = "";
            } else {
                document.getElementById("seaTR").style.display = "none";
                document.getElementById("airTR1").style.display = "none";
                document.getElementById("airTR2").style.display = "none";
            }
        }

        function shipModeSetting() {
            var idx;
            var mode;

            if (document.myform.document.getElementById("SH_MODE") != null) {
                idx = document.myform.document.getElementById("SH_MODE").selectedIndex;
                mode = document.myform.document.getElementById("SH_MODE").options[idx].text;
            } else {
                mode = document.myform.document.getElementById("DL_SH_MODE").value;
            }
            
            if (mode == "SEA") {
                document.getElementById("seaTR").style.display = "";
                document.getElementById("airTR1").style.display = "none";
                document.getElementById("airTR2").style.display = "none";
            } else if (mode == "AIR") {
                document.getElementById("seaTR").style.display = "none";
                document.getElementById("airTR1").style.display = "";
                document.getElementById("airTR2").style.display = "";
            } else {
                document.getElementById("seaTR").style.display = "none";
                document.getElementById("airTR1").style.display = "none";
                document.getElementById("airTR2").style.display = "none";
            }
        }

        function VendorLookUp() {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "vendLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_VEND");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "<%=H_SH_VEND_NO.clientID %>" + "|, " + "<%=SH_VEND_NO.ClientID %>" + "|, " + "<%=H_SH_VEND_NAME.ClientID %>" + "|1, " + "<%=SH_VEND_NAME.ClientID %>" + "|1");
            setInterfaceDataToForm(document.hiddenForm, "sc", "<%=STORER_CODE.selectedvalue %>");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "vendLookUp";
            document.hiddenForm.submit();
        }
        
    </script>
    </div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="viewModeSetting(document.myform);shipModeSetting();" >
    <br />
    <form  id="myform"  runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <input type="hidden" name="moduleAction" value=""/>
    <asp:HiddenField ID="itemList" runat ="server" />
    <asp:HiddenField ID="seqList" runat ="server" />    
    <asp:HiddenField ID="H_SH_VEND_NO" runat="server" />
    <asp:HiddenField ID="H_SH_VEND_NAME" runat="server" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td colspan="6" class="TITLE">
                    <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0"></asp:Table>
                </td>
            </tr>
            <asp:HiddenField ID="lheader" runat ="server" />
            <tr>
                <td colspan="6" class="menuTD">
                <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr>
                    <td class="menuTD" align="left">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <asp:Button ID="CancelBtn" Text ="Cancel Shipment" CssClass="all_button" runat ="server" />
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                    <%If Session("pagemode") <> "N" Then%>             
                        <asp:Button ID="btnShip" Text="Shipped" CssClass="all_button" runat="server" />
                    <% End If%>
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
                        <asp:Label ID="lbl_SH_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="SH_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack="true"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_DATE" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                       <asp:TextBox ID="SH_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" CssClass="REQUIRED"></asp:TextBox>
                     <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="SH_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
            </tr>               
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_CUST_CODE" runat="server" /></font>
                </td>
                <td>
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                    <tr>
                    <td width="50px" runat="server"><asp:DropDownList ID="CUST_CODE" runat="server" MaxLength="20" 
                        AutoPostBack="True" CssClass="REQUIRED"></asp:DropDownList></td>
                    <td><font size="2"></font></td>
                    </tr>
                    </table>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_VEND_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="SH_VEND_NO" runat="server" ReadOnly="true" CssClass="REQUIRED" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="btnVnd_LookUp" runat="server" ImageUrl="../../images/btn_search.gif"
                        onMouseOut="MM_swapImgRestore()" onMousedown="MM_swapImage('btnVnd_LookUp','','../../images/btn_search_over.gif',1)"></asp:ImageButton>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_CUST_NAME" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                    <asp:textbox ID="CUST_NAME" runat="server" ReadOnly="true" 
                        BackColor="#dddddd" Width="200px" /></font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_VEND_NAME" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="SH_VEND_NAME" runat="server" Width="200px" ReadOnly="true" BackColor="#dddddd" ></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_MODE" runat="server" /></font>
                </td>
                <td width="150px" runat="server">     
                     <asp:DropDownList ID="SH_MODE" name="SH_MODE" runat="server" CssClass="REQUIRED" onChange="chgShipMode(this.value);"></asp:DropDownList>     
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_LC_BENEF" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                       <asp:TextBox ID="SH_LC_BENEF" runat="server" Width ="431px" MaxLength="100"></asp:TextBox>   
                </td>             
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_ETD" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="SH_ETD" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                        <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                        <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="SH_ETD" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_ETA" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="SH_ETA" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                        <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                        <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="SH_ETA" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" /> 
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_BOOKING_REF" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="SH_BOOKING_REF" runat="server" Width ="163px" MaxLength="50"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_PORT_LOAD" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="SH_PORT_LOAD" runat="server" MaxLength="50" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_PORT_DISCH" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:TextBox ID="SH_PORT_DISCH" runat="server" MaxLength="50" Width="130px"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_FINAL_DEST" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="SH_FINAL_DEST" runat="server" Width ="165px" MaxLength="50"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_LC_NO" runat="server" /></font>
                </td>
                <td width="150px">     
                     <asp:TextBox ID="SH_LC_NO" runat="server" Width ="163px" MaxLength="100"></asp:TextBox>     
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_BILL_DESC" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                       <asp:TextBox ID="SH_BILL_DESC" runat="server" Width ="349px" MaxLength="100"></asp:TextBox>   
                </td>             
            </tr>
            <tr id="seaTR" style="display: none">
             <td class="LabelTD" colspan="0" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_SH_BL" runat="server" /></font>
            </td>
            <td>     
                 <asp:TextBox ID="SH_BL" runat="server" Width ="163px" MaxLength="100"></asp:TextBox>     
            </td>  
            <td class="LabelTD" nowrap>
                <font size="2">
                    <asp:Label ID="lbl_SH_SEA_VESSEL" runat="server" /></font>
            </td>
            <td colspan="3">
                <font size="2">
                   <asp:TextBox ID="SH_SEA_VESSEL" runat="server" MaxLength="100"></asp:TextBox>   
            </td>            
            </tr>
            <tr id="airTR1" style="display: none">
                 <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_SH_AIR_FLIGHT" runat="server" /></font>
                    </td>
                    <td colspan="5">
                        <font size="2">
                           <asp:TextBox ID="SH_AIR_FLIGHT" runat="server" MaxLength="100"></asp:TextBox>   
                </td>        
            </tr>
            <tr id="airTR2" style="display: none">
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_AIR_AWB" runat="server" /></font>
                </td>
                <td width="150px">     
                     <asp:TextBox ID="SH_AIR_AWB" runat="server" Width ="163px" MaxLength="100"></asp:TextBox>     
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_AIR_MAWB" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                       <asp:TextBox ID="SH_AIR_MAWB" runat="server" MaxLength="100"></asp:TextBox>   
                </td>    
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_AIR_HAWB" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                       <asp:TextBox ID="SH_AIR_HAWB" runat="server" MaxLength="100"></asp:TextBox>   
                </td>         
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_SH_TARRIFF_CODE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <font size="2">
                        <asp:TextBox ID="SH_TARRIFF_CODE" runat="server" MaxLength="50" Width="165px"></asp:TextBox>
                    </font>
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
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select DO" CssClass="all_button" UseSubmitBehavior ="false" />
                </td>
            </tr>
            <tr>
                <td colspan="6">
                                             <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="sd_SEQ" HorizontalAlign="Left">
                           <RowStyle CssClass="GV" />
                        <Columns>
                            <asp:TemplateField HeaderText="No.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_disp_seq" runat="server" Font-Size="11px" style="text-align:center" Width="30px" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Tracking No.">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_track_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_pallet_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Carton No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_carton_no" runat="server" Font-Size="11px" Width="70px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Code">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_itm_code" runat="server" Font-Size="11px" Width="70px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_pack_key" runat="server" Font-Size="11px" Width="60px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="130px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_itm_desc" runat="server" Font-Size="11px" Width="130px" MaxLength="250"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack No.">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_pack_no" runat="server" Font-Size="11px" Width="60px" MaxLength="10"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pack Type">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_pack_type" runat="server" Font-Size="11px" Width="60px" MaxLength="10"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_qty" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="UOM">
                                <ControlStyle Width="60px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="sd_uom" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                              <asp:TemplateField HeaderText="PCS/ UOM">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_pcs_uom" runat="server" Font-Size="11px" 
                                        style="text-align:right" Width="40px" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total PCS">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_totpcs" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total Wgt">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_tot_wgt" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total CBM">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_tot_cbm" runat="server" Font-Size="11px" Width="40px" style="text-align:right" onkeypress="return maskKey(event);"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="130px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="sd_rem" runat="server" Width="130px" Font-Size="10px"></asp:TextBox>
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
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
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