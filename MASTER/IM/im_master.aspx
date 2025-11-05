<%@ Page Language="VB" AutoEventWireup="false" CodeFile="im_master.aspx.vb" Inherits="MASTER_IM_im_master" ValidateRequest="false" MaintainScrollPositionOnPostBack = "true"  %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

<script language="javascript">
    function DisableDeleteButton() {
        var grp = document.getElementsByName("btnDelete");
        var count;

        count = grp.length;
         
        if (count == 1) {
           grp[0].disabled = true;
        }
    }

    function openDetail(val, val2) {
        window.open("altVendorItem.aspx?r=" + val + "&vnd_code=" + val2, "altIM", "menubar=no,scrollbars=yes,resizable=yes,width=1000,height=820,left=5,top=10");

    }

    function maskKey(objEvent) {
        var iKeyCode;
        
        iKeyCode = objEvent.keyCode;

        if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
        
        return false;
    }
            
    function OpenItemLbls() {
            
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "itemStr", document.getElementById("itm_code").innerHTML + "||" + document.getElementById("pack_key").value);
        //setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("storer_code").value);
        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("storer_code").value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Viewstate("storer_code")%>");
        }
        setInterfaceDataToForm(document.hiddenForm, "lblType", "ITEM"); 
        
        //setInterfaceDataToForm(document.hiddenForm, "LabelSize", document.getElementById("LabelSize").value);    
        setInterfaceDataToForm(document.hiddenForm, "LabelSize", 'N');    
        document.hiddenForm.action = "../../REPORT/ITM_LABEL/item_label.aspx";
        document.hiddenForm.target = "ItemLbls";
        document.hiddenForm.submit();
    }

    function OpenItemLbls2() {
            
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "itemStr", document.getElementById("itm_code").innerHTML + "||" + document.getElementById("pack_key").value);
        //setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("storer_code").value);
        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("storer_code").value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Viewstate("storer_code")%>");
        }
        setInterfaceDataToForm(document.hiddenForm, "lblType", "ITEMS"); 
        
        //setInterfaceDataToForm(document.hiddenForm, "LabelSize", document.getElementById("LabelSize").value);    
        setInterfaceDataToForm(document.hiddenForm, "LabelSize", 'N');   

        document.hiddenForm.action = "../../REPORT/ITM_LABEL/item_label.aspx";
        document.hiddenForm.target = "ItemLbls";
        document.hiddenForm.submit();
    }

    function OpenItemLblsC() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "itemStr", document.getElementById("itm_code").innerHTML + "||" + document.getElementById("pack_key").value);
        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("storer_code").value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Viewstate("storer_code")%>");
        }
        setInterfaceDataToForm(document.hiddenForm, "lblType", "ITEM");
        //setInterfaceDataToForm(document.hiddenForm, "LabelSize", document.getElementById("LabelSize").value);    
        document.hiddenForm.action = "../../REPORT/ITM_LABEL/item_labelCable.aspx";
        document.hiddenForm.target = "ItemLbls";
        document.hiddenForm.submit();
    }

    function LocLookUp(wh_val, lb_id, hd_id, wh_id) {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");

            setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            if (wh_id)
                setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_id + "||WH");
            else
                setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
            document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
            document.hiddenForm.target = "locLookUp";
            document.hiddenForm.submit();        
    }

    function OpenWH(itm_code,pk_key) {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "WHSPEC", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=820,height=600,left=5,top=15");
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("storer_code").value);
        setInterfaceDataToForm(document.hiddenForm, "ITM_CODE", itm_code);
        setInterfaceDataToForm(document.hiddenForm, "pack_key", pk_key);
        

        document.hiddenForm.action = "wh_spec.aspx";
        document.hiddenForm.target = "WHSPEC";
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

    function goToImg(flag) {

        //alert("Alert");
        removeAllElementFromForm(document.hiddenForm);
        window.open("", "NewImage", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=' + screen.width + ',height=' + screen.height").focus();

        setInterfaceDataToForm(document.hiddenForm, "FLAG", flag);
        
        document.hiddenForm.action = "../../ShowImg.aspx";
        document.hiddenForm.target = "NewImage";
        document.hiddenForm.submit();
    }

    function GetFileSize(fileid, size) {
       try {
           var fileSize = 0;
                 //for IE
           if ($.browser.msie && parseInt($.browser.version, 10) < 10) {
               //before making an object of ActiveXObject, 
               //please make sure ActiveX is enabled in your IE browser
               //  var objFSO = new ActiveXObject("Scripting.FileSystemObject"); var filePath = $("#" + fileid)[0].value;
               //  var objFile = objFSO.getFile(filePath);
               //  var fileSize = objFile.size; //size in kb
               //  fileSize = fileSize / 1048576; //size in mb
               return true;
            }
             //for FF, Safari, Opeara and Others
           else {
               if ($("#" + fileid)[0].files[0]) {
                   fileSize = $("#" + fileid)[0].files[0].size //size in b
                   fileSize = fileSize / 1024; //size in kb
                   //alert("Uploaded File Size is" + fileSize + "MB");

                   if (fileSize > size) {
                       alert("File Size Cannot Larger than " + size + "MB");
                       return false;
                   }
                   else {
                       return true;
                   }
               }
               return true;
              }
            }
      catch (e) {
              alert("Error is :" + e);
              return false;
           }
      }
      
</script>
<style type="text/css">
select[disabled]
 { 
     border: solid 1px Silver; 
     background-color: White;
     color:Red;  
 }
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
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="javascript:DisableDeleteButton();">
    <br />
    <form name="myform" id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
            <tr>
                <td class="TITLE" colspan="6">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b>
                                    <asp:Label ID="lheader" runat="server" /></b>
                            </td>
                            <td>
                                <a class="link" href="../../IMPORTEXPORT.aspx" target="_blank">IMPORT</a>
                               
                            </td>
                            <td width="100%" class="TITLE">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <asp:Button ID="BtnCopy" runat="server" Text="Add new Pack Key" CssClass="all_button" />
                    <asp:DropDownList runat="server" ID="LabelSize" Visible="false">
                        <asp:ListItem Text="Normal" Value="N" />
                        <asp:ListItem Text="Small" Value="S" />
                        <asp:ListItem Text="A4" Value="A4" />
                    </asp:DropDownList>
                    <asp:Button ID="btnPrint" runat="server" Text="Print Item Label" CssClass="all_button" />

                    <asp:Button ID="btnPrints" runat="server" Text="Print Item Labels" CssClass="all_button" Visible = "false" />

                    <asp:Button ID="btnWh" runat="server" Text="Warehouse Spec." CssClass="all_button" />
                    <asp:Button ID="btnAttach" runat="server" text="Upload/ Download" CssClass="all_button" />
                    <asp:UpdatePanel runat="server" ID="WSDSUDP" RenderMode="Inline">
                        <ContentTemplate>
                            <asp:Button ID="btnMSDS" runat="server" text="MSDS" CssClass="all_button" Visible="false" />
                        </ContentTemplate>
                    </asp:UpdatePanel>
                   <asp:button runat="server" id="btnMsdsUp" text="MSDS" CssClass="all_button" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE")%>'"
                        class="all_button" /></td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_code" runat="server" /></font>
                </td>
                <td >
                    <asp:Label ID="itm_code" runat="server" Font-Size="10" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_pack_key" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="pack_key" runat="server" Width ="177px" MaxLength="20" Text="1"></asp:TextBox>
                </td>
                <td class="LabelTD" width="150px" style="width: 75px">
                    <font size="2">
                        <asp:Label ID="lbl_itm_status" runat="server" /></font>
                </td>
                <td width="150px" style="width: 75px">
                    <asp:Label ID="itm_status" runat="server" Width="157px"></asp:Label>
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_storer_code" runat="server" /></font>
                </td>
                <td runat="server">
                    <asp:DropDownList ID="storer_code" AutoPostBack="true" runat="server" ></asp:DropDownList>
                </td>
                <td class="LabelTD"  nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_UOM" runat="server" /></font>
                </td>
                <td width="150px" >
                    <asp:DropDownList runat="server" ID="itm_UOM" CssClass="REQUIRED" />
                </td>
                <td class="LabelTD"  nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_pcs_per_uom" runat="server" /></font>
                </td>
                <td width="150px" >
                    <asp:TextBox ID="itm_pcs_per_uom" runat="server" Width="80px" cssclass="REQUIRED" onkeypress="return maskKey(event);" MaxLength="9" text="1"></asp:TextBox>
                </td>
            </tr>
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_qty_of_unit" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="itm_qty_of_unit" runat="server" Width="159px" MaxLength="50"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_UOM2" runat="server" text="UOM2"/></font>
                </td>
                <td width="150px" >
                    <asp:DropDownList runat="server" ID="ITM_UOM2"/>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_QTY2" runat="server" text="Qty2"/></font>
                </td>
                <td>
                    <asp:TextBox ID="ITM_QTY2" runat="server" Width="159px" MaxLength="50"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_sku_no" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="itm_sku_no" runat="server" Width="159px" MaxLength="100"></asp:TextBox>
                </td>

            <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_proj_no" runat="server" Text="Project No." />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="proj_no" runat="server" Width ="159px" MaxLength="50" />
                </td>                
            </tr>
             <tr>Qty of Unit (Report Use):
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_name" runat="server" />
                    </font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="itm_name" runat="server" Width="650px" MaxLength="100"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_desc" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="itm_desc" runat="server" Height="78px" Width="652px" 
                        TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>      
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_ref_price" runat="server" Width="100px" /></font>
                </td>
                <td >
                    <asp:TextBox ID="itm_ref_price" runat="server" MaxLength="9" onkeypress="return maskKey(event);"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_price" runat="server" Width="100px" /></font>
                </td>
                <td >
                    <asp:TextBox ID="itm_price" runat="server" MaxLength="9" onkeypress="return maskKey(event);"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_ref_curr" runat="server" Width="100px" /></font>
                </td>
                <td >
                    <asp:TextBox ID="itm_ref_curr" runat="server" MaxLength="10"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_MFG" runat="server" Text="MFG. Part No" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_MFG" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_DRAWING_NO" runat="server" Text="SK Code" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_DRAWING_NO" runat="server" Width ="159px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_pref_wh" runat="server" Width="100px" /></font>
                </td>
                <td id="Td1" runat="server">
                    <asp:DropDownList ID="itm_pref_wh" runat="server" CssClass="REQUIRED" />
                </td>
                <td  nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_pref_loc" runat="server" Visible=false /></font>
                </td>
                <td>
                    <asp:Label ID="dsp_itm_pref_loc"  width="100px" runat="server" Font-Size="11px" />
                    <asp:hiddenfield ID="itm_pref_loc" runat="server" />
                    <asp:Image ID="Image_App_Loc_LookUp"  Visible=false onclick="LocLookUp(document.getElementById('itm_pref_wh').value, 'dsp_itm_pref_loc', 'itm_pref_loc', 'itm_pref_wh')" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand;" align="absmiddle" />
                </td>
                <td  nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_pref_loc2" runat="server" Visible=false text="2nd Pref. Location" /></font>
                </td>
                <td>
                    <asp:Label ID="dsp_itm_pref_loc2"  width="100px" runat="server" Font-Size="11px" />
                    <asp:hiddenfield ID="itm_pref_loc2" runat="server" />
                    <asp:Image ID="Image1"  Visible=false onclick="LocLookUp(document.getElementById('itm_pref_wh').value, 'dsp_itm_pref_loc2', 'itm_pref_loc2','')" runat="server" ImageUrl="../../images/btn_search.gif" onMouseOut="MM_swapImgRestore()" style="border-width:0px;cursor:hand;" align="absmiddle" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_STORES_CLASS" runat="server" Text="Stores Class" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_STORES_CLASS" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_EMB" runat="server" Text="EMB (E/M/B)" />:</font>
                </td>
                <td width="150px">
                    <asp:DropDownList ID="ITM_EMB" runat="server">
                        <asp:ListItem Value="" Text="SELECT"></asp:ListItem>
                        <asp:ListItem Value="E" Text="E"></asp:ListItem>
                        <asp:ListItem Value="M" Text="M"></asp:ListItem>
                        <asp:ListItem Value="B" Text="B"></asp:ListItem>
                    </asp:DropDownList>
                </td>
                 <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_HAZARD_CODE" runat="server" Text="Hazard Code" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_HAZARD_CODE" runat="server" Width ="159px" MaxLength="100" />
                </td>          
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="Lbl_itm_cc" runat="server" Text="Cycle Count Y/N" />:
                    </font>
                </td>
                <td Class="REQUIRED">
                    <asp:CheckBox runat="server" ID="ITM_CC" Checked="true" Text="Yes" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_CC_DATE" runat="server" Width="100px" Text="Last CC Date:"  /></font>
                </td>
                <td colspan="3">
                    <asp:Label runat="server" ID="ITM_CC_DATE" />
                </td>
            </tr>
            <tr>
                 <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_SERIAL_NO_YN" runat="server" Text="Serial Flag (Y/N)" />:</font>
                </td>
                <td width="150px">
                    <asp:CheckBox runat="server" id="ITM_SERIAL_NO_YN" Text="YES" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="Label1" runat="server" Text="Inspection (Y/N)" />:</font>
                </td>
                <td width="150px" colspan="3">
                        <asp:CheckBox runat="server" id="ITM_INSP_YN" Text="YES" />
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_CRITICAL_YN" runat="server" Text="Critical Item (VC/C/NC)" /></font>
                </td>
                <td width="150px">
                    <asp:DropDownList ID="ITM_CRITICAL_YN" runat="server">
                        <asp:ListItem Value="" Text="SELECT"></asp:ListItem>
                        <asp:ListItem Value="VC" Text="VC"></asp:ListItem>
                        <asp:ListItem Value="C" Text="C"></asp:ListItem>
                        <asp:ListItem Value="NC" Text="NC"></asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_SHELF_LIFE" runat="server" Text="Shelf Life" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="ITM_SHELF_LIFE" runat="server" MaxLength="10" /> DAYS
                </td>                
            </tr>
             <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_INT_ORDER_QTY" runat="server" Text="Initial Order Quantity" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_INT_ORDER_QTY" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_MAX_STOCK" runat="server" Text="Min Rec Shelf Life" />:</font>
                </td>
                <td width="150px" colspan="3">
                    <asp:TextBox ID="ITM_MAX_STOCK" runat="server" Width ="159px" MaxLength="100" />
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_ROP_APL" runat="server" Text="ROP (Apleichau)" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_ROP_APL" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_ROP_LAM" runat="server" Text="ROP (Lamma)" />:</font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="ITM_ROP_LAM" runat="server" Width ="159px" MaxLength="100" />
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_ROP_CABLE" runat="server" Text="ROP (Cable Warehouse)" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_ROP_CABLE" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_ROP_NP" runat="server" Text="ROP (North Point)" />:</font>
                </td>
                <td width="150px" colspan="3">
                    <asp:TextBox ID="ITM_ROP_NP" runat="server" Width ="159px" MaxLength="100" />
                </td>
            </tr>
             <tr>
                <td class="auto-style1" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_ORO_YN" runat="server" Text="ORO (Y/N)" />:</font>
                </td>
                <td width="150px" class="auto-style2">
                    <asp:DropDownList runat="server" ID="ITM_ORO_YN" CssClass="REQUIRED">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="Y" Value="Y" Selected="True" />
                        <asp:ListItem Text="N" Value="N" />
                    </asp:DropDownList>
                </td>
                <td class="auto-style1" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_REPLACEMENT_COST" runat="server" Text="Replacement Cost" />:</font>
                </td>
                <td width="150px" colspan="3" class="auto-style2">
                   <asp:TextBox runat="server" ID="ITM_REPLACEMENT_COST" MaxLength="17" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD">
                    <asp:Label runat="server" id="lbl_ITM_MANUFACTURER" Text="Manufacturer:" />
                </td>
                <td colspan="5">
                    <asp:TextBox runat="server" ID="ITM_MANUFACTURER" MaxLength="100" width="300px" /> 
                </td>
            </tr>                   
            <asp:HiddenField ID="itm_vend_itm_no" runat="server" />
            <asp:HiddenField ID="ori_itm_vend_itm_no" runat="server" />
              <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_cat" runat="server" /></font>
                </td>
                <td >
                    <asp:DropDownList ID="itm_cat" runat="server" MaxLength="20" />                  
                    <%-- <asp:TextBox ID="itm_cat" runat="server" Width="159px"></asp:TextBox>--%>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_type" runat="server" />
                      </font>
                </td>
                <td width="150px" colspan="3">
                    <asp:DropDownList ID="itm_type" runat="server"  CssClass="REQUIRED">
                        <asp:ListItem value="">SELECT</asp:ListItem>
                        <asp:ListItem Value="OTHER">Other</asp:ListItem>
                    </asp:DropDownList>
                    <asp:TextBox ID="ITM_PROD_GROUP" runat="server" Width ="159px" MaxLength="20" visible="false" />
                </td>
            </tr>            
            <tr>                
                <td class="LabelTD"  nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_barcode" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="itm_barcode" runat="server" Width ="159px" MaxLength="30"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_parent" runat="server" /></font>
                </td>
                <td >
                    <asp:TextBox ID="itm_parent" runat="server" MaxLength="20" Width="159px"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_KIT_QTY" runat="server" /></font>
                </td>
                <td >
                    <asp:TextBox ID="ITM_KIT_QTY" runat="server" MaxLength="20" Width="159px"></asp:TextBox>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_model" runat="server" /></font>
                </td>
                <td colspan="1">
                    <asp:TextBox ID="itm_model" runat="server" Width="346px" MaxLength="50"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_series" runat="server" /></font>
                </td>
                <td colspan="5" nowrap>
                    <asp:TextBox ID="itm_series_no" runat="server" Width="100px" MaxLength="50"></asp:TextBox>
                    <asp:TextBox ID="itm_series" runat="server" Width="550px" MaxLength="500"></asp:TextBox>
                </td>
            </tr>                       
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_SIZE" runat="server" Text="Size" /></font>:
                </td>
                <td colspan="2">
                    <asp:TextBox ID="ITM_SIZE" runat="server" MaxLength="50" Width="159px" />
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_COLOR_CODE" runat="server" Text="Color Code" /></font>:
                </td>    
                <td colspan="2">
                    <asp:TextBox runat="server" ID="ITM_COLOR_CODE" MaxLength="50" Width="159px" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_SERIAL_NO" runat="server" Text="Serial No." /></font>:
                </td>
                <td colspan="2">
                    <asp:TextBox ID="ITM_SERIAL_NO" runat="server" MaxLength="50" Width="159px" />
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PART_NO" runat="server" Text="Part No." /></font>:
                </td>    
                <td colspan="2">
                    <asp:TextBox runat="server" ID="ITM_PART_NO" MaxLength="50" Width="159px" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PROD_NO" runat="server" Text="Product No." /></font>:
                </td>
                <td >
                    <asp:TextBox ID="ITM_PROD_NO" runat="server" MaxLength="50" Width="159px" />
                </td>
                 <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_brand" runat="server" Width="100px" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="itm_brand" runat="server" Width="251px" MaxLength="50"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_remarks" runat="server" /></font>
                </td>
                <td colspan="5" rowspan="2">
                    <asp:TextBox ID="itm_remarks" runat="server" Height="78px" Width="652px" 
                        TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                </td>      
            </tr>
            <tr><td class="LabelTD" nowrap>&nbsp;</td></tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_balance" runat="server" /></font>
                </td>
                <td >
                    <asp:TextBox ID="itm_balance" runat="server" MaxLength="9" onkeypress="return maskKey(event);"></asp:TextBox>
                </td>
                <asp:Panel runat="server" id="pnl_hide1" Visible='false'>
               <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_EXPENSE_CODE" runat="server" Text="Expense Code" /></font>
               </td>
               <td width="150px" colspan="3">
                    <asp:TextBox ID="ITM_EXPENSE_CODE" runat="server" Width ="159px" MaxLength="100"  CssClass="REQUIRED" />
               </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_GP_CODE" runat="server" Text="Group Code" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_GP_CODE" runat="server" Width ="159px" MaxLength="100" CssClass="REQUIRED" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_KEPT_DIV_CODE" runat="server" Text="Kept for Division" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_KEPT_DIV_CODE" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_DIV_CODE" runat="server" Text="Division Code" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_DIV_CODE" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_COMMODITY_CODE" runat="server" Text="Commodity Code" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_COMMODITY_CODE" runat="server" MaxLength="20" CssClass="REQUIRED" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_RESTRICTED_ITEM" runat="server" Text="Restriction Item" />:</font>
                </td>
                <td colspan="5">
                    <asp:DropDownList ID="ITM_RESTRICTED_ITEM" runat="server" CssClass="REQUIRED">
                        <asp:ListItem Value="" Text="SELECT"></asp:ListItem>
                        <asp:ListItem Value="0" Text="0 - None"></asp:ListItem>
                        <asp:ListItem Value="1" Text="1 = General"></asp:ListItem>
                        <asp:ListItem Value="2" Text="2 = T&D"></asp:ListItem>
                        <asp:ListItem Value="3" Text="3 = Generation"></asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td class="LabelTD">
                    <asp:Label runat="server" id="lbl_ITM_RESTRICTED_RMK" Text="Restriction Remark:" />
                </td>
                <td colspan="5">
                    <asp:TextBox runat="server" Rows="3" TextMode="MultiLine" ID="ITM_RESTRICTED_RMK" Columns="120" />
                </td>
                </asp:Panel>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_lot_exists" runat="server" >LOT# Exists?</asp:Label>
                    </font>
                </td>
                <td colspan="5">
                    <asp:RadioButton ID="radLotExist" runat="server" GroupName="LotExist" Text="YES" />
&nbsp;&nbsp;
                    <asp:RadioButton ID="radLotNotExist" runat="server" GroupName="LotExist" Text="NO" />
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_itm_flags" runat="server" Width="100px" /></font>
                </td>
                <td colspan="5">
                    <asp:checkboxList id="itm_flags" runat="server" RepeatDirection="Horizontal" Visible="false">
                    </asp:checkboxList>                    
                    <%--<asp:CheckBox runat="server" id="ITM_STACKABLE_YN" Text="Stackable" />
                    <asp:CheckBox runat="server" id="ITM_SCRAP_YN" Text="Scrap" />
                    <asp:CheckBox runat="server" ID="ITM_NONSTOCK_YN" Text="Non-Stock" />
                    <br />--%>
                    <asp:CheckBox runat="server" ID="ITM_DG_YN" Text="Dangerous Good" />
                    <asp:Label ID="lbl_ITM_DG_CAT" runat="server" Text="- Category" />
                    <asp:TextBox ID="ITM_DG_CAT" runat="server" Width ="159px" MaxLength="100" />
                    <br />
                    <asp:CheckBox runat="server" ID="ITM_REQ_STORE_HUM_YN" Text="required to store at specific relative humidity" />
                    <asp:TextBox ID="ITM_REQ_STORE_HUM_DESC" runat="server" Width ="159px" MaxLength="100" />
                    <br />
                    <asp:CheckBox runat="server" ID="ITM_REQ_STORE_AIRCON_YN" Text="required to store at Air-Con Warehouse" />
                    <br />
                    <asp:CheckBox runat="server" ID="ITM_REQ_STORE_COLD_YN" Text="required to store at Cold Store" />
                    <br />
                    <asp:CheckBox runat="server" ID="ITM_REQ_MSDS_YN" Text="MSDS is required"/>
                    
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_CHE_CLASS" runat="server" Width="100px" /></font>
                </td>
                <td colspan="5">
                    <asp:checkboxList id="ITM_CHE_CLASS" runat="server" RepeatDirection="Horizontal">
                    </asp:checkboxList>
                </td>
            </tr>                        
            <tr>
                <td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="Label2" runat="server" Width="100px" text="Capital/ Revenue" /></font>
                </td>
                <td>
                    <asp:DropDownList runat="server" ID="ITM_CAP_REV">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="Capital" Value="C" />
                        <asp:ListItem Text="Revenue" Value="R" />
                    </asp:DropDownList>
                </td>
                <td class="LabelTD">
                        <font size="2">
                        <asp:Label ID="lbl_ITM_WEIGHT_TYPE" runat="server" Width="100px" text="Weight (Heavy/Light)" /></font>
                </td>
                <td colspan="3">
                    <asp:DropDownList runat="server" ID="ITM_WEIGHT_TYPE" CssClass="REQUIRED">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="Heavy" Value="Heavy" />
                        <asp:ListItem Text="Light" Value="Light" />
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_ITEM_PRICE_CLASS" runat="server" Width="100px" text="Price Value" /></font>
                </td>
                <td>
                    <asp:DropDownList runat="server" ID="ITEM_PRICE_CLASS">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="High" Value="HI" />
                        <asp:ListItem Text="Medium" Value="MID" />
                        <asp:ListItem Text="Low" Value="LOW" />
                    </asp:DropDownList>
                </td>
                <td class="LabelTD">
                        <font size="2">
                        <asp:Label ID="Label3" runat="server" Width="100px" text="Capital Value" /></font>
                </td>
                <td colspan="3">
                    <asp:DropDownList runat="server" ID="ITM_CAP_VALUE">
                        <asp:ListItem Text="SELECT" Value="" />
                        <asp:ListItem Text="High" Value="H" />
                        <asp:ListItem Text="Medium" Value="M" />
                        <asp:ListItem Text="Low" Value="L" />
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_ITM_1ST_INSP" runat="server" Text="First Inspection to be Done <br>(For Reference)" />:</font>
                    </td>
                    <td colspan="5">
                        <asp:TextBox ID="ITM_1ST_INSP" runat="server" Width ="159px" MaxLength="100" />
                    </td>
            </tr> 

            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_RCV_DATE" runat="server" Text="Size" /></font>:
                </td>
                <td colspan="2">
                    <asp:TextBox ID="ITM_PD_RCV_DATE" runat="server" MaxLength="50" Width="159px" /> (DD/MM/YYYY)
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_CONTRACT_NO" runat="server" Text="Color Code" /></font>:
                </td>    
                <td colspan="2">
                    <asp:TextBox runat="server" ID="ITM_PD_CONTRACT_NO" MaxLength="50" Width="159px" />
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_ARR_NOTICE_NO" runat="server" Text="Size" /></font>:
                </td>
                <td colspan="2">
                    <asp:TextBox ID="ITM_PD_ARR_NOTICE_NO" runat="server" MaxLength="50" Width="159px" />
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_COND_OF_SPARES" runat="server" Text="Color Code" /></font>:
                </td>    
                <td colspan="2">
                    <asp:TextBox runat="server" ID="ITM_PD_COND_OF_SPARES" MaxLength="50" Width="159px" />
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_ST_1_YEAR" runat="server" Text="Size" /></font>:
                </td>
                <td colspan="2">
                    <asp:TextBox ID="ITM_PD_ST_1_YEAR" runat="server" MaxLength="50" Width="159px" />
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_PD_ST_OVER_1_YEAR" runat="server" Text="Color Code" /></font>:
                </td>    
                <td colspan="2">
                    <asp:TextBox runat="server" ID="ITM_PD_ST_OVER_1_YEAR" MaxLength="50" Width="159px" />
                </td>
            </tr>


            <asp:Panel runat="server" ID="hidePNL2" Visible="false">
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_TI_YTD" runat="server" Text="Total Issue - YTD" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_TI_YTD" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_TI_LYR" runat="server" Text="Total Issue - LYR" />:</font>
                </td>
                <td width="150px" colspan="3">
                    <asp:TextBox ID="ITM_TI_LYR" runat="server" Width ="159px" MaxLength="100" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_TI_PY2" runat="server" Text="Total Issue PY2" />:</font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="ITM_TI_PY2" runat="server" Width ="159px" MaxLength="100" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ITM_TI_PY3" runat="server" Text="Total Issue PY2" />:</font>
                </td>
                <td width="150px" colspan="3">
                    <asp:TextBox ID="ITM_TI_PY3" runat="server" Width ="159px" MaxLength="100" />
                </td>
            </tr>
            </asp:Panel>    
            <tr runat="server" id="tempTR" visible="false">
                <td class="LabelTD" nowrap>
                    <font size="2">
                    <asp:Label runat="server" Text="Temperature Y/N?" ID="lbl_ITM_TEMP_YN" />:
                    </font>
                </td>
                <td>
                    <asp:CheckBox runat="server" id="ITM_TEMP_YN" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                    <asp:Label runat="server" Text="From" ID="lbl_ITM_TEMP_FR" />:
                    </font>
                </td>
                <td>
                    <font size="2">
                        <asp:textbox runat="server" ID="ITM_TEMP_FR" width="50" MaxLength="20" />°C
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label runat="server" ID="lbl_ITM_TEMP_TO" Text="To:" />
                    </font>
                </td>
                <td>
                    <asp:textbox runat="server" ID="ITM_TEMP_TO" Width="50" MaxLength="20" />°C
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <br />
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
            <asp:Panel runat="server" ID="oldImage" Visible="false">
            <tr>
            <td colspan="6">
               <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td class="LabelTD" nowrap width="80px">
                                <asp:Label ID="lbl_img" runat="server" Text="Image 1:" />
                            </td>
                            <td Width="300px" >
                                <asp:FileUpload ID="itm_picture1_upload" runat="server" Width="300px" Font-Size="12px"  />
                                <asp:Literal ID="itm_picture1_upload_lit" runat="server" Visible="false"></asp:Literal>
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td class="LabelTD" nowrap width="80px">
                                <asp:Label ID="lbl_img0" runat="server" Text="Image 2:" />
                            </td>
                            <td>
                                <asp:FileUpload ID="itm_picture2_upload" runat="server" Width="300px" Font-Size="12px"  />
                                <asp:Literal ID="itm_picture2_upload_lit" runat="server" Visible="false"></asp:Literal>
                            </td>
                        </tr>
                        <tr>
                            <td nowrap width="80px">&nbsp;</td>
                            <td Width="300px" >
                                <asp:LinkButton ID="itm_picture1_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td nowrap width="80px">&nbsp;</td>
                            <td>
                                <asp:LinkButton ID="itm_picture2_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                            </td>
                        </tr>
                         <tr runat="server" id="pic_tr" visible="false">
                            <td nowrap width="80px">&nbsp;</td>
                            <td>
                                <asp:Label ID="preview_pic1" Font-Size="12px" Font-Italic="True" runat="server" 
                                    Text="Preview: (Click to view actual size)" Width="301px" Visible="False" />
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
                                <asp:LinkButton ID="itm_picture1_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                            </td>
                                 
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td style="padding-bottom: 1px;">
                                <asp:LinkButton ID="itm_picture2_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                            </td>
                        </tr>
                        <tr>
                            <td>&nbsp;</td>
                            <td valign="top">
                                <asp:HyperLink ID="itm_picture1" runat="server" />
                            </td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td valign="top">
                                <asp:HyperLink ID="itm_picture2" runat="server" />
                            </td>
                        </tr>
              </table>
            </td>
            </tr>
            </asp:Panel>
            <tr>
                <td colspan="6">
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td class="LabelTD" nowrap width="80px">
                                <asp:Label ID="lbl_ITM_PHOTO1" runat="server" Text="Image " />
                                <asp:UpdatePanel runat="server" ID="UDP_IMG_SEQ1" RenderMode="Inline">
                                    <ContentTemplate>
                                        <asp:DropDownList runat="server" ID="IMG_SEQ1" Width="40px" AutoPostBack="true">
                                            <asp:ListItem Value="1" Text="1" Selected="True" />
                                            <asp:ListItem Value="2" Text="2" />
                                            <asp:ListItem Value="3" Text="3" />
                                        </asp:DropDownList>    
                                        <asp:HiddenField runat="server" ID="PRE_IMG_SEQ1" Value = "1" />
                                    </ContentTemplate>
                                </asp:UpdatePanel>                                
                            </td>
                            <td Width="300px" >
                                <asp:FileUpload ID="ITM_PHOTO1" runat="server" Width="300px" Font-Size="12px"  />                                
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td class="LabelTD" nowrap width="80px">
                                <asp:Label ID="lbl_ITM_PHOTO2" runat="server" Text="Image" />
                                <asp:UpdatePanel runat="server" ID="UDP_IMG_SEQ2" RenderMode="Inline">
                                    <ContentTemplate>
                                        <asp:DropDownList runat="server" ID="IMG_SEQ2" Width="40px" AutoPostBack="true">
                                            <asp:ListItem Value="1" Text="1"  />
                                            <asp:ListItem Value="2" Text="2" Selected="True" />
                                            <asp:ListItem Value="3" Text="3" />
                                        </asp:DropDownList>
                                        <asp:HiddenField runat="server" ID="PRE_IMG_SEQ2" Value = "2" />
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                            </td>
                            <td>
                                <asp:FileUpload ID="ITM_PHOTO2" runat="server" Width="300px" Font-Size="12px"  />                                
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td class="LabelTD" nowrap width="80px">
                                <asp:Label ID="lbl_ITM_PHOTO3" runat="server" Text="Image" />
                                <asp:UpdatePanel runat="server" ID="UDP_IMG_SEQ3" RenderMode="Inline">
                                    <ContentTemplate>
                                        <asp:DropDownList runat="server" ID="IMG_SEQ3" Width="40px" AutoPostBack="true">
                                            <asp:ListItem Value="1" Text="1" />
                                            <asp:ListItem Value="2" Text="2" />
                                            <asp:ListItem Value="3" Text="3" Selected="True" />
                                        </asp:DropDownList>
                                        <asp:HiddenField runat="server" ID="PRE_IMG_SEQ3" Value = "3" />
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                            </td>
                            <td>
                                <asp:FileUpload ID="ITM_PHOTO3" runat="server" Width="300px" Font-Size="12px"  />                                
                            </td>
                        </tr>                      
                        <tr runat="server" id="PhotoTR" visible="false">
                            <td nowrap width="80px">&nbsp;</td>
                            <td>
                                <asp:Label ID="pre_ITM_PHOTO1" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="250px" Visible="False" />
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td nowrap width="80px">&nbsp;</td>
                            <td>    
                                <asp:Label ID="pre_ITM_PHOTO2" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="250px" Visible="false" />
                            </td>
                            <td style="width:80px">&nbsp;</td>
                            <td nowrap width="80px">&nbsp;</td>
                            <td>    
                                <asp:Label ID="pre_ITM_PHOTO3" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="250px" Visible="false" />
                            </td>
                        </tr>
                        
                                <tr runat="server" id="photoImgTR" visible="false">
                                    <td nowrap width="80px">&nbsp;</td>
                                    <td><asp:UpdatePanel runat="server" ID="UDP_IMGBTN1" RenderMode="Inline">
                                        <ContentTemplate>
                                            <asp:ImageButton runat="server" ID="img_ITM_PHOTO1" style="height: auto; width: auto; max-width: 150px; max-height: 150px;" Visible="false" />
                                        </ContentTemplate>
                                        </asp:UpdatePanel>                        
                                    </td>
                                    <td style="width:80px">&nbsp;</td>
                                    <td nowrap width="80px">&nbsp;</td>
                                    <td>
                                    <asp:UpdatePanel runat="server" ID="UDP_IMGBTN2" RenderMode="Inline">
                                        <ContentTemplate>
                                            <asp:ImageButton runat="server" ID="img_ITM_PHOTO2" style="height: auto; width: auto; max-width: 150px; max-height: 150px;" Visible="false" />
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                    </td>
                                    <td style="width:80px">&nbsp;</td>
                                    <td nowrap width="80px">&nbsp;</td>
                                    <td>
                                    <asp:UpdatePanel runat="server" ID="UDP_IMGBTN3" RenderMode="Inline">
                                        <ContentTemplate>
                                            <asp:ImageButton runat="server" ID="img_ITM_PHOTO3" style="height: auto; width: auto; max-width: 150px; max-height: 150px;" Visible="false" />
                                        </ContentTemplate>
                                    </asp:UpdatePanel>
                                    </td>
                                </tr>
                            
                        <tr runat="server" id="photoDelTR" visible="false">
                            <td>&nbsp;</td>
                             <td style="padding-bottom: 1px;">
                                <asp:LinkButton ID="del_ITM_PHOTO1" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                            </td>
                                 
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td style="padding-bottom: 1px;">
                                <asp:LinkButton ID="del_ITM_PHOTO2" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                            </td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td style="padding-bottom: 1px;">
                                <asp:LinkButton ID="del_ITM_PHOTO3" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                            </td>
                        </tr>                        
              </table>
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
                <td colspan="6">
                    <hr style="width: 100%" />
                </td>
            </tr>
            <tr>
                <td class="TITLE" colspan="6">
                    <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%" class="TITLE">
                                <b><asp:Label ID="lbl_VendorList" runat="server" /></b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left">
                        <Columns>
                            <asp:TemplateField HeaderText="Vendor Code">
                                <ItemTemplate>
                                    <asp:HyperLink ID="vnd_code" runat="server"></asp:HyperLink>
                                </ItemTemplate>
                                <ItemStyle Font-Size="11px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Vendor Name">
                                <ControlStyle Width="250px" />
                                <HeaderStyle HorizontalAlign="Left" Width="250px"  />
                                <ItemTemplate>
                                    <asp:Label ID="vnd_name" runat="server" Font-Size="11px" ></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>   
                            <asp:TemplateField HeaderText="Vendor's Item Code">
                                <HeaderStyle HorizontalAlign="Left" Width="200px"  />
                                <ItemTemplate>
                                    <asp:Label ID="alv_vnd_itmcode" runat="server" Width="200px" Font-Size="10px"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Date">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="alv_date_added" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>  
                            <asp:TemplateField ControlStyle-Width="50px">
                                <ItemTemplate>
                                    <asp:Button ID="btnChgStatus" name="btnChgStatus" runat="server" Height="22px" Font-Size="11px"
                                        CommandName="Update" Text="Active" CssClass="all_button" Font-Bold="false" />
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
                        <HeaderStyle CssClass="TITLE" Font-Size="12px" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
			<table width="100%">
			<tr><td class="menuTD">
<asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
<input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"<% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=<%=Session("PAGE_SESSION_MENU_CODE") %>'" class="all_button" />
			</td><td align="right" class="menuTD"><asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button"  runat ="server" /></td>
			</table>
                </td>
            </tr>
        </table>

       <asp:UpdatePanel runat="server" ID="UDMSDSP">
            <ContentTemplate>
                <asp:Panel ID="pnlMSDS" runat="server" CssClass="modalPopup" style="display: none">
                     <table width="800px">
                         <tr class="TITLE">
                             <td colspan="2" class="TITLE">
                                 <asp:Panel runat="Server" ID="PanelDrag" Style="cursor: move;">
                                   <font size="2">WSDS:</font> </asp:Panel></td>
                         </tr>
                         <tr>
                             <td align="right" colspan="2">
                                 <asp:Button runat="server" ID="btnPUpload" CssClass="all_button" text="Upload" />&nbsp; <asp:Button runat="server" ID="btnPCancel"  CssClass="all_button" Text="Close" />                                 
                             </td>
                         </tr>
                         <tr>
                             <td class="LabelTD" width="100px">
                                <asp:Label runat="server" ID="lbl_APPROV" Text="MSDS Photo:" />
                             </td>
                             <td>
                                <asp:FileUpload runat="server" ID="ITM_MSDS" />
                             </td>
                         </tr>
                         <tr>
                             <td colspan="2" runat="server" id="MSDS_TR" visible="false" align="center" >
                                <asp:Imagebutton runat="server" ID="MSDS_IMG"  OnClientClick="goToImg('');return false;" style="height: auto; width: auto; max-width: 150px; max-height: 150px;" /> <br /><br />
                                <asp:Button runat="server" ID="btnDelMSDS" Text="Delete Photo" CssClass="all_button" />
                             </td>                            
                         </tr>                         
                     </table> 
                    </asp:Panel>            
                    <asp:ModalPopupExtender ID="btnMSDS_ModalPopupExtender" runat="server" 
                                DynamicServicePath="" 
                                Enabled="True" 
                                TargetControlID="dummy" 
                                PopupControlID="pnlMSDS"
                                BackgroundCssClass="modalBackground"
                                DropShadow="true"         
                                CancelControlID="btnPCancel"
                                RepositionMode="RepositionOnWindowResize"
                                PopupDragHandleControlID="PanelDrag"
                                X="50"
                                y="50"
                                 >
                    </asp:ModalPopupExtender>     
                    <asp:HiddenField ID="dummy" runat="server" />    
        </ContentTemplate>
            <Triggers>
                <asp:PostBackTrigger runat="server" ControlID="btnPUpload" />
            </Triggers>
        </asp:UpdatePanel>
    </div>
     <asp:HiddenField ID="editMode" runat="server" />
    </form>
    <form id="hiddenForm" name="hiddenForm" method="post" />
</body>
</html>
