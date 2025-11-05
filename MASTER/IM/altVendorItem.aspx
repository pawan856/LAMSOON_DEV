<%@ Page Language="VB" AutoEventWireup="false" CodeFile="altVendorItem.aspx.vb" Inherits="MASTER_IM_altVendorMaster" ValidateRequest="false" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
  <title></title>

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
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
  
   function VendorLookUp(hv_code,v_code,hv_name, v_name) {
       removeAllElementFromForm(document.hiddenForm);

       window.open("", "vendLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");

       setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_VEND");
       setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
       setInterfaceDataToForm(document.hiddenForm, "pItemList", hv_code + "|, " + v_code + "|, " + hv_name + "|1, "  + v_name + "|1");
       setInterfaceDataToForm(document.hiddenForm, "sc",  document.getElementById("storer_code").innerHTML);
       document.hiddenForm.action = "../../cms_search.aspx";
       document.hiddenForm.target = "vendLookUp";
       document.hiddenForm.submit();
   }
  
   function maskKey(objEvent) {
       var iKeyCode;
       iKeyCode = objEvent.keyCode;
       if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
       return false;
   }

   function updtCBM() {
       //alert('test');
       document.getElementById("aitm_cbm").value = fixDecimal(document.getElementById("aitm_length").value * document.getElementById("aitm_width").value * document.getElementById("aitm_hight").value / 1000000, 3);
       //alert(document.getElementById("aitm_cbm").value);
   }

</script>

<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css">
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <br />
  <form method="post" name="myform" id="myform" runat="server">
  <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />    
  <div id="div1">
    <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
      <tr>
        <td class="TITLE" colspan="7">
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
        <td colspan="7" class="menuTD">
          <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
          <input id="btnBack2" type="button"  value="Close" onclick="Javascript:window.close();"
            class="all_button" /></td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
            <font size="2">
            <asp:Label ID="lbl_itm_code" runat="server" /></font>
        </td>
        <td width="250px">
          <asp:Label ID="itm_code" runat="server" Width="203px"></asp:Label>
        </td>
        <td class="LabelTD" colspan="0" nowrap>
          <font size="2">
            <asp:Label ID="lbl_pack_key" runat="server" /></font>
        </td>
        <td>
        <asp:Label ID="pack_key" runat="server" Width ="106px"></asp:Label>
        </td>
        <td class="LabelTD" width="150px" style="width: 75px">
          <font size="2">
            <asp:Label ID="lbl_aitm_status" runat="server" /></font>
        </td>
        <td width="150px" colspan="2" style="width: 75px">
          <asp:Label ID="aitm_status" runat="server" Width="157px"></asp:Label>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
            <font size="2">
            <asp:Label ID="lbl_storer_code" runat="server" /></font>
        </td>
        <td >
          <asp:Label ID="storer_code" runat="server" ></asp:Label>
        </td>
        <td class="LabelTD" colspan="0" nowrap>
            <font size="2">
            <asp:Label ID="lbl_alv_date_added" runat="server" /></font>
        </td>
        <td colspan="4">
          <asp:TextBox ID="alv_date_added" runat="server" Width="80px" MaxLength="10"></asp:TextBox>
             <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="alv_date_added" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
        </td>
      </tr>
       <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_vnd_code" runat="server" /></font>
        </td>
        <td colspan="3" >
             <asp:TextBox ID="vnd_code" runat="server" Width="105px" CssClass="REQUIRED"></asp:TextBox>
            <asp:ImageButton ID="btnVnd_LookUp" runat="server" ImageUrl="../../images/btn_search.gif"
             onMouseOut="MM_swapImgRestore()" onMousedown="MM_swapImage('btnVnd_LookUp','','../../images/btn_search_over.gif',1)"></asp:ImageButton>
            <asp:TextBox ID="vnd_name" runat="server" Width="286px" ReadOnly="true"></asp:TextBox>
        </td>
         <td class="LabelTD" colspan="2">
          <font size="2">
            <asp:Label ID="lbl_alv_vnd_itmcode" runat="server" Width="150px" /></font>
              </td>
        <td>
          <asp:TextBox ID="alv_vnd_itmcode" runat="server" Width="142px" MaxLength="20"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_name" runat="server" />
          </font>
        </td>
        <td colspan="6">
          <asp:TextBox ID="aitm_name" runat="server" Width="423px" MaxLength="100"></asp:TextBox>
        </td>
        <!--<td class="LabelTD" colspan="2">
          <font size="2"> -->
            <asp:Label ID="lbl_aitm_uom" runat="server" Visible="false" />
        <!--      </font>
        </td>
        <td> -->
             <asp:DropDownList ID="aitm_uom" runat="server" visible="false" />
        <!--</td> -->
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_desc" runat="server" /></font>
        </td>
        <td colspan="6" rowspan="2">
          <asp:TextBox ID="aitm_desc" runat="server" Height="78px" Width="652px" 
            TextMode="MultiLine" MaxLength="200"></asp:TextBox>
        </td>   
      </tr>
       <tr><td class="LabelTD" nowrap>&nbsp;</td></tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_spec" runat="server" /></font>
        </td>
        <td colspan="6" rowspan="2">
          <asp:TextBox ID="aitm_spec" runat="server" Height="78px" Width="652px" 
            TextMode="MultiLine" MaxLength="200"></asp:TextBox>
        </td>   
      </tr>
      <tr><td class="LabelTD" nowrap>&nbsp;</td></tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_remarks" runat="server" /></font>
        </td>
        <td colspan="6" rowspan="2">
          <asp:TextBox ID="aitm_remarks" runat="server" Height="78px" Width="652px" 
            TextMode="MultiLine" MaxLength="200"></asp:TextBox>
        </td>   
      </tr>
      <tr><td class="LabelTD" nowrap>&nbsp;</td></tr>
       <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_pcs_per_pack" runat="server" Width="100px" /></font>
        </td>
        <td>
          <asp:TextBox ID="aitm_pcs_per_pack" runat="server" Width="80px" cssclass="REQUIRED" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_qty_per_ctn" runat="server" Width="100px" /></font>
        </td>
        <td colspan="4">
          <asp:TextBox ID="aitm_qty_per_ctn" runat="server" Width="80px" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_AITM_ORIGIN" runat="server" Width="100px" /></font>
        </td>
        <td colspan="6" >
           <asp:TextBox ID="AITM_ORIGIN" runat="server" MaxLength="30" Width="200px"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_sizes" runat="server" Width="100px" /></font>
        </td>
        <td>
          <asp:TextBox ID="aitm_length" runat="server" Width="105px" onkeypress="return maskKey(event);" onchange="updtCBM();" MaxLength="9"></asp:TextBox>&nbsp;<font size="2"><asp:Label 
                ID="lbl_aitm_length" Text="Length" runat="server" Width="45px" /></font> 
        </td>
        <td colspan="2" width="300px">
           <asp:TextBox ID="aitm_width" runat="server" onkeypress="return maskKey(event);" onchange="updtCBM();" MaxLength="9"></asp:TextBox>&nbsp;<font size="2"><asp:Label ID="lbl_aitm_width" Text="Width" runat="server" Width="100px" /></font>
        </td> 
        <td colspan="3" width="300px">
          <asp:TextBox ID="aitm_hight" runat="server" onkeypress="return maskKey(event);" onchange="updtCBM();" MaxLength="9"></asp:TextBox>&nbsp;<font size="2"><asp:Label ID="lbl_aitm_hight" Text="Height" runat="server" Width="100px" /></font>
        </td>
      </tr>
       <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_cbm" runat="server" Width="100px" /></font>
        </td>
        <td >
          <asp:TextBox ID="aitm_cbm" runat="server" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_vol" runat="server" /></font>
        </td>
        <td colspan="4" >
          <asp:TextBox ID="aitm_vol" runat="server" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_AITM_NET_WEIGHT" runat="server" Width="100px" /></font>
        </td>
        <td >
          <asp:TextBox ID="AITM_NET_WEIGHT" runat="server" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_AITM_GROSS_WEIGHT" runat="server" /></font>
        </td>
        <td colspan="4" >
          <asp:TextBox ID="AITM_GROSS_WEIGHT" runat="server" onkeypress="return maskKey(event);" MaxLength="9"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_pref_wh" runat="server" Width="100px" /></font>
        </td>
        <td runat="server">
           <asp:DropDownList ID="aitm_pref_wh" runat="server"></asp:DropDownList>
        </td>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_pref_loc" runat="server" /></font>
        </td>
        <td colspan="4" >
          <asp:TextBox ID="aitm_pref_loc" runat="server" MaxLength="20"></asp:TextBox>
        </td>
      </tr>
       <tr>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_ref_price" runat="server" Width="100px" /></font>
        </td>
        <td >
          <asp:TextBox ID="aitm_ref_price" runat="server" onkeypress="return maskKey(event);" MaxLength="9" ReadOnly="true"></asp:TextBox>
        </td>
        <td class="LabelTD" nowrap>
          <font size="2">
            <asp:Label ID="lbl_aitm_ref_curr" runat="server" Width="100px" /></font>
        </td>
        <td colspan="4" >
          <asp:TextBox ID="aitm_ref_curr" runat="server" MaxLength="10" ReadOnly="true"></asp:TextBox>
        </td>
      </tr>
      <tr>
        <td colspan="7">
          <br />
        </td>
      </tr>
      <tr>
        <td class="TITLE" colspan="7">
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
      <td colspan="7">
        <table border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td class="LabelTD" nowrap width="80px">
                <asp:Label ID="lbl_img" runat="server" Text="Image 1:" />
              </td>
              <td Width="300px" >
                <asp:FileUpload ID="aitm_picture1_upload" runat="server" Width="300px" Font-Size="12px" />
                <asp:Literal ID="aitm_picture1_upload_lit" runat="server" Visible="false"></asp:Literal>
              </td>
              <td style="width:80px">&nbsp;</td>
              <td class="LabelTD" nowrap width="80px">
                <asp:Label ID="lbl_img0" runat="server" Text="Image 2:" />
              </td>
              <td>
                <asp:FileUpload ID="aitm_picture2_upload" runat="server" Width="300px" Font-Size="12px" />
                <asp:Literal ID="aitm_picture2_upload_lit" runat="server" Visible="false"></asp:Literal>
              </td>
            </tr>
             <tr>
                <td nowrap width="80px">&nbsp;</td>
                <td Width="300px" >
                    <asp:LinkButton ID="aitm_picture1_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                </td>
                <td style="width:80px">&nbsp;</td>
                <td nowrap width="80px">&nbsp;</td>
                <td>
                    <asp:LinkButton ID="aitm_picture2_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                </td>
            </tr>
             <tr runat="server" id="pic_tr" visible="false">
              <td nowrap width="80px">&nbsp;</td>
              <td>
                <asp:Label ID="preview_pic1" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="250px" Visible="false" />
              </td>
              <td style="width:80px">&nbsp;</td>
              <td nowrap width="80px">&nbsp;</td>
              <td>
                <asp:Label ID="preview_pic2" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="250px" Visible="false" />
              </td>
            </tr>
            <tr>
              <td nowrap width="80px">&nbsp;</td>
              <td>&nbsp;</td>
              <td style="width:80px">&nbsp;</td>
              <td nowrap width="80px">&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
            <tr runat="server" id="remove_tr" visible="false">
              <td>&nbsp;</td>
              <td style="padding-bottom: 1px;">
                <asp:LinkButton ID="aitm_picture1_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
              </td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td style="padding-bottom: 1px;">
                <asp:LinkButton ID="aitm_picture2_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
              </td>
              </tr>
            <tr>
              <td>&nbsp;</td>
              <td valign="top">
                <asp:HyperLink ID="aitm_picture1" runat="server" />
              </td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
              <td valign="top">
                <asp:HyperLink ID="aitm_picture2" runat="server" />
              </td>
            </tr>
       </table>
      </td>
      </tr>
      <tr>
        <td colspan="7">
          <hr style="width: 100%" />
        </td>
      </tr>
      <tr>
        <td colspan="7" align="left" width="100%">
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
        <td colspan="7" class="menuTD">
          <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
          <input id="btnBack" type="button" value="Close" onclick="Javascript:window.close();" class="all_button" />
        </td>
      </tr>
    </table>
  </div>
    <asp:HiddenField ID="h_vnd_code" runat="server" />
    <asp:HiddenField ID="h_vnd_name" runat="server" />
  </form>
  <form name="hiddenForm" id="hiddenForm" method="post"></form>
</body>
</html>
