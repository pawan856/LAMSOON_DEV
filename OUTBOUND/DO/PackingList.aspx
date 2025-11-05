<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PackingList.aspx.vb" Inherits="OUTBOUND_DO_PackingList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Packing List</title>
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

function LocLookUp_old(wh_val, lb_id, hd_id, wh_id, lb_fl, hd_fl, lb_ar, hd_ar, lb_rk, hd_rk, lb_bn, hd_bn) {
    if(document.piform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id + ", " + wh_id + "||WH, " + lb_fl + "|L|FL, " + hd_fl + "||FL, " + lb_ar + "|L|AR, " + hd_ar + "||AR, " + lb_rk + "|L|RK, " + hd_rk + "||RK, " + lb_bn + "|L|BN, " + hd_bn + "||BN");
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function LocLookUp(itm_code, pack_key, pallet_no) {
    if(document.piform.editMode.value!="V")
    {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "pForm", "piform");
        setInterfaceDataToForm(document.hiddenForm, "IMP_CODE", document.piform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.piform.STORER_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "ITM_CODE", itm_code);
        setInterfaceDataToForm(document.hiddenForm, "PACK_KEY", pack_key);
        setInterfaceDataToForm(document.hiddenForm, "PALLET_NO", pallet_no);
        setInterfaceDataToForm(document.hiddenForm, "ref_no", opener.myform.DO_CUS_REF_NO.value);
        document.hiddenForm.action = "./PickListLookup.aspx";
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

function PrintPL() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintPACKING", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1180,height=700,left=5,top=15");
 
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById('STORER_CODE').value);
        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE').value);
        document.hiddenForm.action = "packingList_printout.aspx";
        document.hiddenForm.target = "doPrintPACKING";
        document.hiddenForm.submit();
    }

    function splitCarton(carStr) {
        var carArray = carStr.split(",");
        var num = 0;
        var count = 0;

        while (num < carArray.length) {
            if (carArray[num] != '') {
                if (carArray[num] != ',') {
                    var idx = carArray[num].indexOf('-');

                    //alert(carArray[num]);

                    if (idx > 0) {
                        var toList = carArray[num].split('-');
                        var fromNo = -1;
                        var toNo = -99;

                        var lnum = 0;

                        while (lnum < toList.length) {
                            if (lnum > 1) {
                                break;
                            } else {

                                if (isNaN(parseInt(toList[lnum], 10)) == false) {
                                    if (lnum == 0) fromNo = toList[lnum];
                                    if (lnum == 1) toNo = toList[lnum];
                                }

                                if (fromNo != toNo && fromNo < toNo) {
                                    for (i = fromNo; i <= toNo; i++) {
                                        count += 1
                                    }
                                }
                                lnum += 1;
                            }
                        }

                    } else {
                        count += 1
                    }
                }
            }

            num += 1;
        }

        return count;
    }

    function calTotalQty(carID, packID, qtyID, totqtyID) {
        var carQty = 0;
        var qty = document.getElementById(qtyID).value;
        var packSize = document.getElementById(packID).value;
        //alert(document.getElementById(packID).value);
        //alert(document.getElementById(qtyID).value);
        if (qty == "") qty = 0;
        if (packSize == "") packSize = 0;

        //carQty = splitCarton(document.getElementById(carID).value);

        //if (carQty == 0) carQty = 1;

        //document.getElementById(totqtyID).innerHTML = carQty * packSize * qty;

        document.getElementById(totqtyID).value = packSize * qty;
       
        
    }

    function calPickQty(TotalID,qtyID, PackID,oriPack,oriQty) {

        var total = document.getElementById(TotalID).value;
        var qty = document.getElementById(qtyID).value;


        if (total != "" && qty != "") {
           if (total==0 || qty == 0 ){
           
           }
           else{
                if (total % qty != 0) {
                    alert('Pack Size entered cannot be completely divided by Total Size.');
                    document.getElementById(qtyID).value = oriPack;
                    document.getElementById(PackID).value = oriQty;
                      }
                else {
                    document.getElementById(PackID).value = total / qty;
                     }
                }
        } else {
       
        document.getElementById(qtyID).value = 0;
        }
        //alert(carQty);
    }
    
 function calPickQty2(TotalID, qtyID, PackID, oriTotal) {
        var total = document.getElementById(TotalID).value;
        var qty = document.getElementById(qtyID).value;


      if (total != "" && qty != "") {
                    if (total==0 || qty == 0 ){

          }
          else{
                 if (total % qty != 0) {
                     alert('Total Number entered cannot be completely divided by Pack Size.');
//                     alert(total + '/' + qty);
                    document.getElementById(TotalID).value = oriTotal;
                      }
               else {
                   document.getElementById(PackID).value = total / qty;
                     }
                 }
         }
          else {
              document.getElementById(TotalID).value=0;           
          }
         //alert(carQty);
    }

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:DisableDeleteButton();">
<form name="piform" id="piform" runat="server">
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
            <td class="menuTD" align="left">
                <!--<asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" />-->
                <asp:Button ID="saveBtn1" runat="server" Text="OK" CssClass="all_button" />
                <input id="btnClose1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                    <% Elseif Session("gLang") = "C" Then %>value="关闭" <% End If%> onclick="Javascript:window.close();"
                    class="all_button" />
            </td>            
            <td class="menuTD" align="right">
                <asp:Button ID="btnReset" Text="Reset" CssClass="all_button" runat="server" />
                <%--<asp:Button ID="btnPrint" Text="Print" CssClass="all_button" runat="server" OnClientClick="PrintPL();return;"  />--%>
                <input type="button" value="Print" onclick="PrintPL();" class="all_button" />                
            </td>
        </tr>
        </table>
        </td>
    </tr>
    <tr><td  width="5%">&nbsp;</td>
        <td class="DtlLabel" Width="80px" style="text-align: right">
        <asp:TextBox runat="server" ID="DO_PACK_LABEL_1" MaxLength="20" Width="80px"/>
        </td>
        <td width="300px"><asp:TextBox runat="server" ID="DO_PACK_LABEL_QTY_1" MaxLength="12" Width="80px"/></td>      
        <td class="DtlLabel" Width="80px" style="text-align: right">
        <asp:TextBox runat="server" ID="DO_PACK_LABEL_2" MaxLength="20" Width="80px"/>
        </td>
        <td width="300px"><asp:TextBox runat="server" ID="DO_PACK_LABEL_QTY_2" MaxLength="12"  Width="80px"/></td>
        <td class="DtlLabel" Width="80px" style="text-align: right">
        <asp:TextBox runat="server" ID="DO_PACK_LABEL_3" MaxLength="20" Width="80px"/>
        </td>
        <td width="300px"><asp:TextBox runat="server" ID="DO_PACK_LABEL_QTY_3" MaxLength="12"  Width="80px"/></td>
        <td  width="5%">&nbsp;</td>
    </tr>
    <tr>
        <td colspan="8">
            <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" 
                EmptyDataText="No Record Found." CaptionAlign="Top" HorizontalAlign="Left" >
                <Columns>
                    <asp:TemplateField HeaderText="No.">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:Label ID="pad_display_seq" runat="server" Font-Size="11px" style="text-align:center" />
                        </ItemTemplate>
                    </asp:TemplateField>                        
                    <asp:TemplateField HeaderText="Carton No.">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="pad_carton_no" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pallet No.">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="pad_pallet_no" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Lot No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="PAD_REF_NO" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Origin">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="PAD_ORIGIN" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vendor Code">
                        <ControlStyle Width="70px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="pad_vnd_code" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Item Code">
                        <ControlStyle Width="130px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Label ID="pad_itm_code" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Batch No.">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>                        
                            <asp:Label ID="pad_batch_no" runat="server" Font-Size="11px"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Pack Size">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:textbox ID="pad_pack_size" runat="server" MaxLength="20" Font-Size="11px" />
                            <asp:HiddenField id="pad_pack_key" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty. Per Carton">
                        <ControlStyle Width="35px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:textbox ID="PAD_QTY_PER_CTN" runat="server" MaxLength="20" Font-Size="11px" />                            
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Picked Qty">
                        <ControlStyle Width="50px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:TextBox ID="PAD_QTY" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Total Number">
                        <ControlStyle Width="60px" />
                        <HeaderStyle HorizontalAlign="Right" />
                        <ItemTemplate>
                            <asp:textbox ID="itm_total_qty" runat="server" Font-Size="11px" maxlength="12" style="text-align:right" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="N.W">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Textbox ID="PAD_NET_WEIGHT" runat="server" Font-Size="11px"></asp:Textbox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="G.W">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:Textbox ID="PAD_GROSS_WEIGHT" runat="server" Font-Size="11px"></asp:Textbox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="L">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="PAD_LENGTH" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="W">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="PAD_WIDTH" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="H">
                        <ControlStyle Width="30px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="pad_height" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Min. Packing">
                        <ControlStyle Width="40px" />
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemTemplate>
                            <asp:TextBox ID="pad_min_packing" runat="server" Font-Size="11px" maxlength="12" style="text-align:right"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Packed By">
                        <ControlStyle Width="80px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="pad_pack_by" runat="server" Font-Size="11px" maxlength="20"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField ControlStyle-Width="50px">
                        <ItemTemplate>
                            <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px"
                                CommandName="SplitItem" Text="Split" CssClass="all_button" Font-Bold="false" UseSubmitBehavior="false" />
                        </ItemTemplate>
                        <ControlStyle Width="50px"></ControlStyle>
                        <HeaderStyle Width="50px" />
                    </asp:TemplateField>     
                    <asp:TemplateField ControlStyle-Width="50px">
                        <ItemTemplate>
                            <asp:Button ID="btnDelete" name="btnDelete" runat="server" Height="22px" Font-Size="11px"
                                CommandName="Delete" Text="Delete" CssClass="all_button" Font-Bold="false" UseSubmitBehavior="false" />
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
<asp:HiddenField ID="DO_CODE" runat="server" />
<asp:HiddenField ID="editMode" runat="server" />
<asp:HiddenField ID="moduleAction" runat="server" />
</form>
<form name="hiddenForm" id="hiddenForm" method="post" />
</body>
</html>
