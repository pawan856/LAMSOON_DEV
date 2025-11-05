<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CustomerMaster.aspx.vb" ValidateRequest="false"  Inherits="MASTER_CM_CustomerMaster" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Customer Master</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

    <script language="javascript" src="../../js/validation.js"></script>
    <script language="javascript" src="../../js/JS_Calendar.js"></script>
    <script language="javascript" src="../../js/listUtil.js"></script>
    <script language="javascript" src="../../js/formatUtil.js"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js"></script>
    <script language="javascript">

        function DisableDeleteButton() {
            var grp = document.getElementsByName("btnDelete");
            var count;
            count = grp.length;
            if (count == 1) {
                grp[0].disabled = true;
            }
        }

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
        function maskNumOnlyLOT(objEvent) {
            var iKeyCode;
            iKeyCode = objEvent.keyCode;
            if (iKeyCode === 48 || iKeyCode === 49 || iKeyCode == 50 || iKeyCode == 51) {
                return true;
            }
            alert("Valid MAX LOTS (0,1,2,3)")
            return false;
        }
        function maskNumOnly(objEvent) {
            var iKeyCode;
            iKeyCode = objEvent.keyCode;
            if (iKeyCode >= 48 && iKeyCode <= 57) return true;
            return false;
        }

    </script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form id="Form1" runat="server">
        <%-- <form method="post" name="myform" id="myform" onsubmit="" runat="server">--%>
        <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
        <input type="hidden" name="moduleAction" value="" />
        <div id="div1">
            <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%">
                <tr>
                    <td class="TITLE" colspan="8">
                        <table border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="100%" class="TITLE">
                                    <b>
                                        <asp:Label ID="lheader" runat="server" /></b>
                                </td>
                                <td width="100%" class="TITLE"></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="8" class="menuTD">
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" /></td>
                </tr>
                <tr>
                    <td class="LabelTD" width="20%" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CODE" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_CODE" runat="server" Text="" Font-Size="10" MaxLength="20"></asp:TextBox>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_LOGO_L" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:FileUpload ID="CUS_LOGO_L_UPLOAD" Visible="false" runat="server"
                            Width="350px" />
                        <asp:Literal ID="CUS_LOGO_L_UPLOAD_lit" runat="server" Visible="false"></asp:Literal>
                        <asp:LinkButton ID="CUS_LOGO_L_EDIT" runat="server" Width="100px"
                            Visible="false" Height="16px" Style="float: left; vertical-align: bottom">Change Image</asp:LinkButton>
                        <asp:LinkButton ID="CUS_LOGO_L_REMOVE" runat="server" Width="96px"
                            Visible="false"
                            OnClientClick="return confirm(&quot;Are you sure to Remove Large Logo?&quot;);"
                            Height="16px" Style="float: left; vertical-align: bottom">Remove</asp:LinkButton>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" width="20%" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_STORER_CODE" runat="server" />
                        </font>
                    </td>
                    <td colspan="3" runat="server">
                        <asp:DropDownList ID="STORER_CODE" runat="server"></asp:DropDownList>
                    </td>
                    <td rowspan="5"></td>
                    <td rowspan="5" colspan="3" valign="top">
                        <asp:HyperLink ID="CUS_LOGO_L" runat="server" BorderWidth="1px" BorderStyle="Ridge"></asp:HyperLink>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" width="20%" nowrap>
                        <asp:Label ID="lbl_CUS_STATUS" runat="server" />
                    </td>
                    <td colspan="3" runat="server">
                        <font size="2">
                            <asp:DropDownList ID="CUS_STATUS" runat="server"></asp:DropDownList>
                        </font>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_SHORTNAME" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_SHORTNAME" runat="server" Width="400" MaxLength="20" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_NAME" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_NAME" runat="server" Width="400" MaxLength="100" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_NAME_CH" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_NAME_CH" runat="server" Width="400" MaxLength="100" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap rowspan="3">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_ADDR_DEL" runat="server" />:</font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_ADDR1_DEL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                    <td class="LabelTD" width="20%" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_LOGO_S" runat="server" />
                        </font>
                    </td>
                    <td colspan="3">
                        <asp:FileUpload ID="CUS_LOGO_S_UPLOAD" Visible="false" runat="server"
                            Width="350px" />
                        <asp:Literal ID="CUS_LOGO_S_UPLOAD_lit" runat="server" Visible="false"></asp:Literal>
                        <asp:LinkButton ID="CUS_LOGO_S_EDIT" runat="server" Width="100px"
                            Visible="false" Height="16px" Style="float: left; vertical-align: bottom">Change Image</asp:LinkButton>
                        <asp:LinkButton ID="CUS_LOGO_S_REMOVE" runat="server" Width="96px"
                            Visible="false"
                            OnClientClick="return confirm(&quot;Are you sure to Remove Small Logo?&quot;);"
                            Height="16px" Style="float: left; vertical-align: bottom">Remove</asp:LinkButton>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_ADDR2_DEL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                    <td rowspan="3"></td>
                    <td colspan="3" rowspan="3" valign="top">
                        <asp:HyperLink ID="CUS_LOGO_S" runat="server" BorderWidth="1px" BorderStyle="Ridge"></asp:HyperLink>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_ADDR3_DEL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_COUNTRY_DEL" runat="server" /></font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_COUNTRY_DEL" runat="server" MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_REGION_DEL" runat="server" /></font>
                    </td>
                    <td colspan="7">
                        <asp:TextBox ID="CUS_REGION_DEL" runat="server" MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_AREA_DEL" runat="server" /></font>
                    </td>
                    <td colspan="7">
                        <asp:TextBox ID="CUS_AREA_DEL" runat="server" MaxLength="50" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" rowspan="4">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_ADDR_BILL" runat="server" />:</font>
                    </td>
                </tr>
                <tr>
                    <td colspan="7">
                        <asp:TextBox ID="CUS_ADDR1_BILL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <asp:TextBox ID="CUS_ADDR2_BILL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                    <td>
                        <font size="2">
                            <asp:Label ID="lbl_MAX_LOT" runat="server" >MAX LOTS</asp:Label>
                        </font>
                    </td>
                    <td colspan="2">
                        <asp:TextBox Width="50px" runat="server" MaxLength="1" ID="txtMaxLots" AutoPostBack="True" onkeypress="return maskNumOnlyLOT(event);" />
                        </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <asp:TextBox ID="CUS_ADDR3_BILL" runat="server" Width="400px" MaxLength="100" />
                    </td>
                    <td>
                        <asp:CheckBox ID="chkMPL" Width="150px" runat="server" Text="Min. Production Life" AutoPostBack="true" />
                    </td>

                    <td>
                        <asp:TextBox Width="100%" runat="server" MaxLength="3" ID="txtMPLMonths" Enabled="false" AutoPostBack="True" onkeypress="return maskNumOnly(event);" />
                        <%--<asp:TextBox ID="MIN_PROD_LIFE" runat="server" Width ="85" MaxLength="10" onkeypress="return maskDate(event);" Enabled="false" AutoPostBack="True"></asp:TextBox>
                       <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                        <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" TargetControlID="MIN_PROD_LIFE" PopupButtonID="ImageButton1" Format="dd-MMM-yyyy" />--%>
                    </td>
                    <td>
                        <span style="padding-left: 10px">
                            <asp:Label runat="server" Text="Days" />
                        </span>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_COUNTRY_BILL" runat="server" /></font>
                    </td>
                    <td colspan="4">
                        <asp:TextBox ID="CUS_COUNTRY_BILL" runat="server" MaxLength="3" />
                    </td>
                    <td>
                        <asp:CheckBox ID="chkMSL" Width="150px" runat="server" Text="Min. Self Life" AutoPostBack="True" />
                    </td>

                    <td>
                        <%-- <asp:TextBox ID="MIN_SELF_LIFE" runat="server" Width ="85" MaxLength="10" onkeypress="return maskDate(event);" Enabled="false" AutoPostBack="True"></asp:TextBox>
                       <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                       <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" TargetControlID="MIN_SELF_LIFE" PopupButtonID="ImageButton2" Format="dd-MMM-yyyy"/>--%>

                        <asp:TextBox Width="100%" runat="server" MaxLength="3" ID="txtMSLMonths" Enabled="false" AutoPostBack="True" onkeypress="return maskNumOnly(event);" />
                    </td>

                    <td>
                        <span style="padding-left: 10px">
                            <asp:Label runat="server" Text="Days" />
                        </span>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_REGION_BILL" runat="server" /></font>
                    </td>
                    <td colspan="4">
                        <asp:TextBox ID="CUS_REGION_BILL" runat="server" MaxLength="50" />
                    </td>
                    <td>
                        <asp:CheckBox ID="chkMaxBatches" Width="150px" runat="server" MaxLength="50" Text="Max. Batches" AutoPostBack="True" Visible="False" />
                    </td>
                    <td>
                        <asp:TextBox Width="100%" runat="server" MaxLength="50" ID="txtMaxBatch" Enabled="false" onkeypress="return maskNumOnly(event);" Visible="False" />
                    </td>
                    <td></td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_AREA_BILL" runat="server" /></font>
                    </td>
                    <td colspan="4">
                        <asp:TextBox ID="CUS_AREA_BILL" runat="server" MaxLength="50" />
                    </td>
                    <td colspan="2">
                        <asp:CheckBox ID="chkValidLot" Width="100%" runat="server" MaxLength="50" Text="Lot must not be earlier." />
                    </td>

                    <td></td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap style="border: 2px;">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_PER_GEN" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_PER_GEN" runat="server" MaxLength="50" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_TEL_GEN" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_TEL_GEN" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_DEPT_GEN" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_DEPT_GEN" runat="server" MaxLength="30" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_EMAIL_GEN" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_EMAIL_GEN" runat="server" Width="165px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap style="border: 2px;">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_PER_ACC" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_PER_ACC" runat="server" MaxLength="50" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_TEL_ACC" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_TEL_ACC" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_DEPT_ACC" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_DEPT_ACC" runat="server" MaxLength="30" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_EMAIL_ACC" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_EMAIL_ACC" runat="server" Width="165px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap style="border: 2px;">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_PER_ORD" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_PER_ORD" runat="server" MaxLength="50" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_TEL_ORD" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_TEL_ORD" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_DEPT_ORD" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_DEPT_ORD" runat="server" MaxLength="30" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_EMAIL_ORD" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_EMAIL_ORD" runat="server" Width="165px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap style="border: 2px;">
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_PER_LOG" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_PER_LOG" runat="server" MaxLength="50" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_TEL_LOG" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_TEL_LOG" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_DEPT_LOG" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_DEPT_LOG" runat="server" MaxLength="30" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CONT_EMAIL_LOG" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_CONT_EMAIL_LOG" runat="server" Width="165px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_WEBSITE" runat="server" /></font>
                    </td>
                    <td colspan="7">
                        <asp:TextBox ID="CUS_WEBSITE" runat="server" Width="400px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_FAX" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_FAX" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_MAIN_TEL" runat="server" /></font>
                    </td>
                    <td>
                        <asp:TextBox ID="CUS_MAIN_TEL" runat="server" onkeypress="return maskKey(event);" MaxLength="20" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_MAIN_EMAIL" runat="server" /></font>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="CUS_MAIN_EMAIL" runat="server" Width="165px" MaxLength="30" />
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_CURR" runat="server" /></font>
                    </td>
                    <td runat="server">
                        <asp:DropDownList ID="CUS_CURR" runat="server"></asp:DropDownList>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_PAYTERM" runat="server" /></font>
                    </td>
                    <td colspan="4" runat="server">
                        <asp:DropDownList ID="CUS_PAYTERM" runat="server"></asp:DropDownList>
                    </td>
                    <td colspan="1" runat="server" align="right">
                        <asp:Button ID="btnAddRow" runat="server" Text="Add Row"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_CUS_REM" runat="server" /></font>
                    </td>
                    <td colspan="2">
                        <asp:TextBox ID="CUS_REM" runat="server" TextMode="MultiLine" Width="90%" Height="80" MaxLength="200" />
                    </td>
                    <td colspan="5">
                        <asp:GridView Width="100%" ID="gvRules" runat="server" AutoGenerateColumns="False">
                            <Columns>
                                <asp:TemplateField HeaderText="ITEM CODE">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="gvDdlITEMCODE" runat="server" Width="100px">
                                            <asp:ListItem Value="SELECT"></asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:HiddenField ID="gvHfITEM_CODE" runat="server" Value='<%# Bind("ITEM_CODE") %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MPL FLAG">
                                    <%--<EditItemTemplate>
                                        <asp:CheckBox ID="gvChkMPL" runat="server" Checked='<%#If(Eval("MPL_FLAG").ToString() = "1", True, False) %>'/>
                                    </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="gvChkMPL" runat="server" Checked='<%#If(Eval("MPL_FLAG").ToString() = "1", True, False) %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MIN PROD LIFE">
                                    <%--<EditItemTemplate>
                                        <asp:TextBox ID="gvTxtMPLMonths" runat="server" MaxLength="3" Text='<%# Bind("MIN_PROD_LIFE") %>'></asp:TextBox>
                                    </EditItemTemplate>--%>

                                    <ItemTemplate>
                                        <%--<asp:TextBox ID="MIN_PROD_LIFE" runat="server" Width ="85px" MaxLength="10" Text='<%# Bind("MIN_PROD_LIFE") %>' onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                       <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" TargetControlID="MIN_PROD_LIFE" PopupButtonID="ImageButton3" Format="dd-MMM-yyyy" />--%>
                                        <asp:TextBox ID="gvTxtMPLMonths" runat="server" MaxLength="3" Width ="55px" Text='<%# Bind("MIN_PROD_LIFE") %>'></asp:TextBox>

                                        <span style="padding-left: 5px">
                                            <asp:Label runat="server" Text="Days" />
                                        </span>

                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MSL FLAG">
                                    <%--<EditItemTemplate>
                                         <asp:CheckBox ID="gvChkMPL" runat="server" Checked='<%#If(Eval("MSL_FLAG").ToString() = "1", True, False) %>'/>
                                        </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="gvChkMSL" runat="server" Checked='<%#If(Eval("MSL_FLAG").ToString() = "1", True, False) %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MIN SELF LIFE">
                                    <%--<EditItemTemplate>
                                        <asp:TextBox ID="gvTxtMSLMonths" runat="server" MaxLength="3" Text='<%# Bind("MIN_SELF_LIFE") %>'></asp:TextBox>
                                    </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <%-- <asp:TextBox ID="MIN_SELF_LIFE" runat="server" Width ="85" MaxLength="10" Text='<%# Bind("MIN_SELF_LIFE") %>' onkeypress="return maskDate(event);"></asp:TextBox>
                       <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                       <asp:CalendarExtender ID="CalendarExtender4" runat="server" CssClass="ajax_calendar" TargetControlID="MIN_SELF_LIFE" PopupButtonID="ImageButton4" Format="dd-MMM-yyyy" />--%>

                                        <asp:TextBox ID="gvTxtMSLMonths" runat="server" MaxLength="3" Width ="55px" Text='<%# Bind("MIN_SELF_LIFE") %>'></asp:TextBox>
                                        <span style="padding-left: 5px">
                                            <asp:Label runat="server" Text="Days" />
                                        </span>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MB FLAG">
                                    <%--<EditItemTemplate>
                                        <asp:CheckBox ID="gvChkMB" Checked='<%#If(Eval("MB_FLAG").ToString() = "1", True, False) %>' runat="server" />
                                    </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="gvChkMB" Checked='<%#If(Eval("MB_FLAG").ToString() = "1", True, False) %>' runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MAX BATCHES" Visible="False">
                                    <%--<EditItemTemplate>
                                        <asp:TextBox ID="gvTxtMBMonths" runat="server" MaxLength="3" Checked='<%#If(Eval("MAX_BATCHES").ToString() = "1", True, False) %>'></asp:TextBox>
                                    </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <asp:TextBox ID="gvTxtMBMonths" runat="server" MaxLength="3" Text='<%# Bind("MAX_BATCHES") %>' Width="85" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="LOTS CANNOT BE EARLIER">
                                    <%--<EditItemTemplate>
                                        <asp:CheckBox ID="gvChkValidLot" runat="server" Checked='<%#If(Eval("LOTS_CANNOT_BE_EARLIER").ToString() = "1", True, False) %>'/>
                                    </EditItemTemplate>--%>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="gvChkValidLot" runat="server" Checked='<%#If(Eval("LOTS_CANNOT_BE_EARLIER").ToString() = "1", True, False) %>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="MAX LOTS">
                                    <ItemTemplate>
                                        <asp:TextBox ID="gvTxtMaxLots" runat="server" MaxLength="1" onkeypress="return maskNumOnlyLOT(event);" Text='<%# Bind("MAX_LOTS") %>' Width="85"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:ButtonField CommandName="Remove" Text="Remove" />
                            </Columns>
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td colspan="8" align="left" width="100%">
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
                    <td colspan="8" class="menuTD">
                        <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                            class="all_button" />
                    </td>
                </tr>
            </table>
        </div>
        <br />
        <div id="div2">
        </div>
    </form>
</body>
</html>
