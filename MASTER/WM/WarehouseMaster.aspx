<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WarehouseMaster.aspx.vb" Inherits="MASTER_WM_WarehouseMaster" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <title>Warehouse Master</title>
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

function OpenLocMain() {
    //if (document.myform.editMode.value != "V") {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "locMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1000,height=700,resizable=yes,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.WH_CODE.value);
    document.hiddenForm.action = "../../LOOKUP/locMain.aspx";
    document.hiddenForm.target = "locMain";
    document.hiddenForm.submit();
    //}
}

function OpenDelDtl() {
    //if (document.myform.editMode.value != "V") {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "deldtl", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=600,resizable=yes,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.WH_CODE.value);
    document.hiddenForm.action = "del_dtl.aspx";
    document.hiddenForm.target = "deldtl";
    document.hiddenForm.submit();
    //}
}


function OpenExcelMain() {
    //if (document.myform.editMode.value != "V") {
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "expExcel", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,resizable=yes,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "WH_CODE", document.myform.WH_CODE.value);
    document.hiddenForm.action = "PrintLoc.aspx";
    document.hiddenForm.target = "expExcel";
    document.hiddenForm.submit();
    //}
}

</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <br />
    <form method="post" name="myform" id="myform" onsubmit="" runat="server">
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
                         <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <%If Session("pagemode") <> "N" Then%>             
                        <input id="btnOpenLoc" type="button" <%if Session("gLang") = "E" Then %>value="Loc. Defn."
                        <% Elseif Session("gLang") = "C" Then %>value="New / Change" <% End If%> onclick="Javascript:OpenLocMain();"
                        class="all_button" />
                        <input runat="server" type="button" id="btnExcel" value="Export Loc. code to Excel" onclick="javascript:OpenExcelMain();" class="all_button" />
                        <input type="button" runat="server" id="btnDel" value="Delete Location" onclick="javascript:OpenDelDtl();" class="all_button" />
                        <% End If%>                    
                    </td>
                </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" width="160px" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CODE" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:textbox ID="WH_CODE" runat="server" Text="" Font-Size="10" MaxLength="3"></asp:textbox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" width="160px" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_OWNER" runat="server" /> </font>
                </td>
                <td colspan="1">
                    <asp:textbox ID="WH_OWNER" runat="server" Text="" Font-Size="10" MaxLength="100"></asp:textbox>
                </td>
                <td class="LabelTD" width="160px" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_MAIN_WH" runat="server" /> </font>
                </td>
                <td colspan="5">
                    <asp:textbox ID="WH_MAIN_WH" runat="server" Text="" Font-Size="10" MaxLength="100"></asp:textbox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_NAME" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_NAME" runat="server" Width="400" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_NAME_CH" runat="server" /> </font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_NAME_CH" runat="server" Width="400" MaxLength="100" />
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_ADDR" runat="server" />:</font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_ADDR1" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="WH_ADDR2" runat="server" Width="400px" MaxLength="100" /><br />
                    <asp:TextBox ID="WH_ADDR3" runat="server" Width="400px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_COUNTRY" runat="server"  /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_COUNTRY" runat="server" MaxLength="50" />
                    </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_REGION" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_REGION" runat="server" MaxLength="50" />
                </td>
           </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_TYPE" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:DropDownList ID="WH_TYPE" runat="server" Width="150px">
                    </asp:DropDownList>
                </td>
            </tr>
             <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_PER1" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_PER1" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_TEL1" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_TEL1" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_EMAIL1" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_EMAIL1" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
            </tr>        
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_PER2" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_PER2" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_TEL2" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_TEL2" runat="server" MaxLength="20" onkeypress="return maskTel(event);"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_CONT_EMAIL2" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="WH_CONT_EMAIL2" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
            </tr>   
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_GROSS_AREA" runat="server" /></font>
                </td>
                <td nowrap>
                    <asp:TextBox ID="WH_GROSS_AREA" runat="server" Width="90px" onkeypress="return maskDecimal(this, 9, 2);" /><b></b>m<sup>2</sup></b>&nbsp;
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_GROSS_CBM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="WH_GROSS_CBM" runat="server" Width="90px" onkeypress="return maskDecimal(this, 9, 2);" />
                </td>                
            </tr>     
                   
      <tr>
      <td class="LabelTD" nowrap width="160px"><asp:Label ID="lbl_img" runat="server" Text="Image:" /></td>
      <td colspan="6">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
              <td>
                <br />
                <asp:FileUpload ID="wh_picture_upload" runat="server" Width="400px" Font-Size="12px" />
                  <br />
                <asp:Literal ID="wh_picture_upload_lit" runat="server" Visible="false"></asp:Literal>
              </td>
            </tr>
             <tr>
                <td>
                    <asp:LinkButton ID="wh_picture_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                </td>
            </tr>
             <tr runat="server" id="pic_tr" visible="false">
              <td>
                <asp:Label ID="preview_pic" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="300" Visible="false" />
              </td>
            </tr>
            <tr>
              <td>&nbsp;</td>
            </tr>
            <tr runat="server" id="remove_tr" visible="false">
              <td style="padding-bottom: 1px;">
                <asp:LinkButton ID="wh_picture_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
              </td>
              </tr>
            <tr>
              <td valign="top">
                <asp:HyperLink ID="wh_picture" runat="server" />
              </td>
            </tr>
       </table>
      </td>
      </tr>        
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_WH_REM" runat="server" /></font>
                </td>
                <td colspan="7">
                    <asp:TextBox ID="WH_REM" runat="server" TextMode="MultiLine" Width="400" Height="80" MaxLength="200" />
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
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div id="div2">
    </div>
    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>