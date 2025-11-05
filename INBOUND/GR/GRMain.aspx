<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GRMain.aspx.vb" Inherits="GRMain" ValidateRequest="false" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server"> <title>Good Receiving</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script src="../../js/jquery-1.9.1.min.js"></script>
<script src="../../js/jquery-migrate-1.1.1.min.js"></script> 
<div runat="server" id="DIVSCRIPT">
<script language="javascript">
var grd_batch = [];
var grd_qty = [];
var grd_dc = [];
var grd_lot = [];

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

function maskNumOnly(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if (iKeyCode >= 48 && iKeyCode <= 57) return true;
    return false;
}

function checkKey(objEvent) {
    alert(objEvent.keyCode);
}

function VendorLookUp() {

    if(document.myform.editMode.value!="V")
    {
        document.getElementById('FocusBarYN').value = 'N';
        
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "vendLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=600,height=500,left=5,top=15");
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_VEND");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "dsp_VND_CODE||L, VND_CODE, dsp_VND_NAME|1|L, VND_NAME|1");
        setInterfaceDataToForm(document.hiddenForm, "sc", document.myform.STORER_CODE.value);
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "vendLookUp";
        document.hiddenForm.submit();
    }
}

function LocLookUp(wh_val, lb_id, hd_id) {
    if(document.myform.editMode.value!="V")
    {
        document.getElementById('FocusBarYN').value = 'N';
        
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yesstatus=1,width=800,height=700,left=5,top=15,resizable=yes");
        
        setInterfaceDataToForm(document.hiddenForm, "wh", wh_val);
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
        document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
        document.hiddenForm.target = "locLookUp";
        document.hiddenForm.submit();
    }
}

function OpenPalletLbls() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grPalletLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1150,height=700,left=5,top=15");

    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    document.hiddenForm.action = "GR_LABELS/pallet_labels.aspx";
    document.hiddenForm.target = "grPalletLbls";
    document.hiddenForm.submit();
}

function OpenCartonLbls() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grCartonLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=880,height=800,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    document.hiddenForm.action = "GR_LABELS/carton_labels.aspx";
    document.hiddenForm.target = "grCartonLbls";
    document.hiddenForm.submit();
}

function OpenItemLbls() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    document.hiddenForm.action = "GR_LABELS/item_label_print.aspx";
    document.hiddenForm.target = "grItemLbls";
    document.hiddenForm.submit();
}


function ROLookUp(txt_GR_DOC_NO, GR_DOC_NO, STORER_CODE) {
    if (STORER_CODE == '')
    {
        alert('Please select the Storer first!');
        return false;
    }

    if(document.myform.editMode.value!="V")
    {
        document.getElementById('FocusBarYN').value = 'N';
        
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ROLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=750,left=5,top=15");
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_RO");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedRO()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "ROList|1, ROSeqList|2, qtyList|CC_ITEM_QTY#txt_item_qty");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ROLookUp";
        document.hiddenForm.submit();
    }
}

function OpenGRReport() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grStockReceipt", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=0,width=1200,height=800,left=5,top=15");
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);    
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    setInterfaceDataToForm(document.hiddenForm, "gr_doc_type", document.myform.GR_DOC_TYPE.value);
    document.hiddenForm.action = "GR_STOCK_REP/stock_receipt.aspx";
    document.hiddenForm.target = "grStockReceipt";
    document.hiddenForm.submit();
}


function OpenCargoRCP() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grCargoReceipt", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=0,width=600,height=150");
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);    
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    document.hiddenForm.action = "Cargo_REP/CARGO_REP.aspx";
    document.hiddenForm.target = "grCargoReceipt";
    document.hiddenForm.submit();
}

function OpenPAReport() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "grPAList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=0,width=880,height=850");
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);    
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.myform.GR_CODE_HF.value);
    document.hiddenForm.action = "GR_LIST/paList_print.aspx";
    document.hiddenForm.target = "grPAList";
    document.hiddenForm.submit();
}

function selectedRO()
{   
        document.getElementById('FocusBarYN').value = 'N';
        document.myform.moduleAction.value = "SELECTRO";
        document.myform.submit();
}

function ItemLookUp(STORER_CODE) {
    if (STORER_CODE == '')
    {
        alert('Please select the Storer first!');
        return false;
    }

    if(document.myform.editMode.value!="V")
    {
        document.getElementById('FocusBarYN').value = 'N';
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "ItemLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=500,left=5,top=15");
        
        setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_IM");
        setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
        setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedItem()");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2");
        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "ItemLookUp";
        document.hiddenForm.submit();
    }
}

function selectedItem()
{   
        document.getElementById('FocusBarYN').value = 'N';
        document.myform.moduleAction.value = "SELECTIM";
        document.myform.submit();
}

function checkSB(STORER_CODE) {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "sbLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=1200,height=500,left=5,top=15");
    
    setInterfaceDataToForm(document.hiddenForm, "menu_code", "INQ_001");
    setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
    setInterfaceDataToForm(document.hiddenForm, "ctemp", true);
    setInterfaceDataToForm(document.hiddenForm, "screadonly", true); 

    document.hiddenForm.action = "../../cms_search.aspx";
    document.hiddenForm.target = "sbLookUp";
    document.hiddenForm.submit();
}

function SerialLookUp(i_code, i_name, i_qty) {
    //if (document.myform.editMode.value != "V") {
        document.getElementById('FocusBarYN').value = 'N';
        
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "snLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=1000,height=700,left=5,top=15,resizable=yes");

        setInterfaceDataToForm(document.hiddenForm, "i_code", i_code);
        setInterfaceDataToForm(document.hiddenForm, "i_name", i_name);
        setInterfaceDataToForm(document.hiddenForm, "i_qty", i_qty);
        
        if(document.myform.editMode.value!="V")
        {
            setInterfaceDataToForm(document.hiddenForm, "sc", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "sc", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "gr", document.myform.GR_CODE_HF.value);

        document.hiddenForm.action = "genSerialNo.aspx";
        document.hiddenForm.target = "snLookUp";
        document.hiddenForm.submit();
    //}
}

function inspMain() {

    removeAllElementFromForm(document.hiddenForm);

    window.open("", "inspMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1150,height=600,left=5,top=15");
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);    
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
    setInterfaceDataToForm(document.hiddenForm, "gr_status", document.getElementById('GR_STATUS').innerHTML);
    setInterfaceDataToForm(document.hiddenForm, "gr_code", document.getElementById("GR_CODE_HF").value);
    document.hiddenForm.action = "GRInsp.aspx";
    document.hiddenForm.target = "inspMain";
    document.hiddenForm.submit();
}

function putAwayMain() {
    document.getElementById('FocusBarYN').value = 'N';
    
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "putAwayMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1100,height=800,left=5,top=15");
    
    if (document.myform.editMode.value != "V") {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);    
    } else {
        setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
    }
    
    setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
    setInterfaceDataToForm(document.hiddenForm, "gr_status", document.getElementById('GR_STATUS').innerHTML);
    setInterfaceDataToForm(document.hiddenForm, "gr_doc_no", document.myform.GR_DOC_NO.value);
    setInterfaceDataToForm(document.hiddenForm, "gr_doc_type", document.myform.GR_DOC_TYPE.value);
    document.hiddenForm.action = "PutAway.aspx";
    document.hiddenForm.target = "putAwayMain";
    document.hiddenForm.submit();
}

function saveok()
{
    document.getElementById('FocusBarYN').value = 'N';
    //document.myform.moduleAction.value = "SAVEOK";
    //document.myform.submit();

    submitAction("SAVEOK", "updtPnlPostBack");
}

function submitAction(moduleAction, postBackID) {
    if (document.getElementById("moduleAction").value != moduleAction || moduleAction == "") {
        document.getElementById("moduleAction").value = moduleAction;

        if (postBackID != "")
            __doPostBack(postBackID, '');
        else
            __doPostBack('__Page', '');
    }
}

function copySizeTo(rowIdx) {
    var DataGridObj = document.getElementById('GridView1');
 
    if (<%=ViewState("size_index")%> != -1) 
    {
        DataGridObj.rows[rowIdx].cells[<%=ViewState("size_index")%>].children[0].value = document.myform.C_L.value;
        DataGridObj.rows[rowIdx].cells[<%=ViewState("size_index")%>].children[1].value = document.myform.C_W.value;
        DataGridObj.rows[rowIdx].cells[<%=ViewState("size_index")%>].children[2].value = document.myform.C_H.value;
    }
    
    if (<%=ViewState("cbmIndex")%> != -1) 
    {
        DataGridObj.rows[rowIdx].cells[<%=ViewState("cbmIndex")%>].children[0].value = fixDecimal(document.myform.C_L.value * document.myform.C_W.value * document.myform.C_H.value) / 1000000 * parseFloat(DataGridObj.rows[rowIdx].cells[<%=ViewState("qtyIndex")%>].children[0].value.replace(/,/gi, ""));
    }
    
    if (<%=ViewState("kg_index")%> != -1) 
    {
        DataGridObj.rows[rowIdx].cells[<%=ViewState("kg_index")%>].children[0].value = document.myform.C_KG.value;
    }
   
    //sumAll();

 }

function splitCarton(carStr) {
    var carArray =carStr.split(",");
    var num=0;
    var count=0;
    
    while (num < carArray.length)
     {
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
                            
                            if (fromNo != toNo && fromNo < toNo){
                               for (i=fromNo;i<=toNo;i++) {
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
        
        num+=1;
      }
      
      
    }  
    
     function sumAll() {
        var DataGridObj = document.getElementById('GridView1');

        var tot = 0;
        var qtot = 0;
        var totKg = 0;
        var totCbm = 0;
        var totCt = 0;
        
        var eachVal;
        var qVal;
        var ctVal;
        
        var idx;
        var qIdx;
        var cIdx;
        var kIdx;
        var ctIdx;
        
        idx = <%=ViewState("cIndex")%>;
        qIdx = <%=ViewState("qtyIndex")%>;
        kIdx = <%=ViewState("kgIndex")%>;
        cIdx = <%=ViewState("cbmIndex")%>;
        ctIdx = <%=ViewState("carIndex")%>;

        if (DataGridObj.rows.length > 0) {
            for (var iRow = 1; iRow <= DataGridObj.rows.length - 1; iRow++) {
                if (iRow == DataGridObj.rows.length -1) {
                             
                    DataGridObj.rows[iRow].cells[idx].align = "Right";
                    DataGridObj.rows[iRow].cells[idx].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[idx].style.fontWeight = "bold";
                    
                    DataGridObj.rows[iRow].cells[qIdx].align = "Right";
                    DataGridObj.rows[iRow].cells[qIdx].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[qIdx].style.fontWeight = "bold";
                    
                    DataGridObj.rows[iRow].cells[kIdx].align = "Right";
                    DataGridObj.rows[iRow].cells[kIdx].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[kIdx].style.fontWeight = "bold";
                    
                    DataGridObj.rows[iRow].cells[cIdx].align = "Right";
                    DataGridObj.rows[iRow].cells[cIdx].style.fontSize = 11;
                    DataGridObj.rows[iRow].cells[cIdx].style.fontWeight = "bold";
                    
                    DataGridObj.rows[iRow].cells[idx].innerText = tot.toFixed(2);
                    //DataGridObj.rows[iRow].cells[qIdx].innerText = qtot.toFixed(0);
                    DataGridObj.rows[iRow].cells[kIdx].innerText = totKg.toFixed(2);
                    DataGridObj.rows[iRow].cells[cIdx].innerText = totCbm.toFixed(2);
                    DataGridObj.rows[iRow].cells[ctIdx].innerText = totCt.toFixed(0);
      
                   
                } else {
                    eachVal = 0;
                    qVal = 0;
                    
                    if (DataGridObj.rows[iRow].cells[idx].children[0].value != "") {
                        eachVal = parseFloat(DataGridObj.rows[iRow].cells[idx].children[0].value.replace(/,/gi, ""));
                        tot = parseFloat(tot) + parseFloat(eachVal);
                    }

                    if (DataGridObj.rows[iRow].cells[ctIdx].children[0].value != "") {
                        ctVal = parseFloat(DataGridObj.rows[iRow].cells[ctIdx].children[0].value.replace(/,/gi, ""));
                        totCt = parseFloat(totCt) + parseFloat(ctVal);
                    }

                    if (DataGridObj.rows[iRow].cells[qIdx].children[0].value != "") {
                        qVal = parseFloat(DataGridObj.rows[iRow].cells[qIdx].children[0].value.replace(/,/gi, ""));
                        qtot = parseFloat(qtot) + parseFloat(qVal);
                        
                        if (qVal > 0) {
                            if (DataGridObj.rows[iRow].cells[kIdx].innerText != "") {
                                totKg = parseFloat(totKg) + parseFloat(DataGridObj.rows[iRow].cells[kIdx].innerText.replace(/,/gi, "").replace(" ", ""));
                            }
                            if (DataGridObj.rows[iRow].cells[cIdx].children[0].value != "") {
                                totCbm = parseFloat(totCbm) + parseFloat(DataGridObj.rows[iRow].cells[cIdx].children[0].value.replace(/,/gi, ""));
                            }               
                        }   
                    }
                }
            }
        } 
    } 
    
    function calOSQty(rcvQtyArray, roQtyObj, osQtyObj) {
        var totRcvQty = 0;
        
        if (rcvQtyArray != null) {
                for (i in rcvQtyArray) {
                    totRcvQty += parseFloat(rcvQtyArray[i].value);
                } 
        }
        
        osQtyObj.value = parseFloat(roQtyObj.value) - totRcvQty;

    }

    function calOSQtyTemp(roQtyObj, rcvQtyObj, osQtyObj) {
        var roQty, rcvQty;

        if (roQtyObj.value)
            roQty = parseFloat(roQtyObj.value);
        else
            roQty = 0;

        if (rcvQtyObj.value)
            rcvQty = parseFloat(rcvQtyObj.value);
        else
            rcvQty = 0;

        osQtyObj.value = roQty - rcvQty;

    }
    
    function calcKg(cID, lID, kgID) {
        var obj = document.getElementById(lID);
        var kgobj = document.getElementById(kgID);
        var qobj = document.getElementById(cID);
        
        if (obj != null) {
            obj.innerHTML = fixDecimal(kgobj.value * qobj.value,2);
        }
    }
    
     function bindcolor(idx) {
        var DataGridObj = document.getElementById('GridView1');

        for (var iCell = 0; iCell < DataGridObj.rows[idx].cells.length; iCell++) {
            if (DataGridObj.rows[idx].cells[iCell].style.backgroundColor == "") {
                DataGridObj.rows[idx].cells[iCell].style.backgroundColor = "#FFCC00";
            } else {
                DataGridObj.rows[idx].cells[iCell].style.backgroundColor = "";
            }
        }
    }
    
     function selectRow(idx) {
        var DataGridObj = document.getElementById('GridView1');

        for (var iRow = 2; iRow < DataGridObj.rows.length; iRow++) {
              for (var iCell = 0; iCell < DataGridObj.rows[iRow].cells.length; iCell++) {
                if (iRow == idx) {
                    DataGridObj.rows[iRow].cells[iCell].style.backgroundColor = "#FFCC00";
                } else {
                    DataGridObj.rows[iRow].cells[iCell].style.backgroundColor = "";
                }
            }
        }
      
    }
    
    function chkBarcode(objEvent) {
        var iKeyCode;
        iKeyCode = objEvent.keyCode;
        if (iKeyCode == 13) {
             PageMethods.scanBarcode(document.getElementById('txtBarCode').value, document.myform.STORER_CODE.value, function(returnValue) {
                                document.getElementById('txtBarCode').focus();
                                var itm = returnValue[0];
                                var batch = returnValue[1];
                                var qty = returnValue[2];
                                var date = returnValue[3];
                                
                                if(returnValue[4]!=''){
                                    document.getElementById('sel_row_idx').value = returnValue[4];
                                    selectRow(returnValue[4]);
                                } else {
                                    
                                }
                                
                                if ( document.getElementById('sel_row_idx').value != "") {
                                        var DataGridObj = document.getElementById('GridView1');
                                        
                                        var idx;
                                        
                                        idx = parseInt( document.getElementById('sel_row_idx').value) - 2;
                                        
                                        if (returnValue[1] != "") document.getElementById(grd_batch[idx]).value = returnValue[1];
                                        if (returnValue[2] != "") {
                                            document.getElementById(grd_qty[idx]).value = returnValue[2];
                                            calOSQty(eval('TM'+document.getElementById(grd_qty[idx].replace('grd_rcv_qty','grd_seq')).value),
                                                        document.getElementById(grd_qty[idx].replace('grd_rcv_qty','grd_po_qty')),
                                                        document.getElementById(grd_qty[idx].replace('grd_rcv_qty','grd_os_qty')))
                                            document.getElementById(grd_qty[idx].replace('grd_rcv_qty','grd_tot_pcs')).value = document.getElementById(grd_qty[idx].replace('grd_rcv_qty','grd_pcs_per_uom')).value * returnValue[2];
                                            sumAll();
                                        }                                               
                                        if (returnValue[3] != "") document.getElementById(grd_dc[idx]).value = returnValue[3];
                                        //if (returnValue[5] != "") document.getElementById(grd_lot[idx]).value = returnValue[5];
                                        
                                }
                                
                                if (returnValue[6] != "") document.getElementById('txtOutput').innerHTML = returnValue[6];
            });
            
            document.getElementById('txtBarCode').value = "";
            document.getElementById('txtBarCode').focus();
            return;
        } else {
            return false;
        }
    }
    
    function FocusBarCode() {
        if (document.getElementById('FocusBarYN').value == 'Y') {
                if(document.myform.editMode.value!="V") {
                    if (document.forms[0].txtBarCode != null) {
                        document.forms[0].txtBarCode.focus();   
                    }
                }            
        } else {
            document.getElementById('FocusBarYN').value = 'N';
        }
    }

    function reloadPage(alertMsg) {
        if (alertMsg)
            alert(alertMsg);

        removeAllElementFromForm(document.hiddenForm);
        setInterfaceDataToForm(document.hiddenForm, 'GR_CODE', document.getElementById("GR_CODE_HF").value);

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("STORER_CODE").value);    
        } else {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
        }
        
        document.hiddenForm.action = './GRMain.aspx';
        document.hiddenForm.target = '_self';
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

    function showSerial(gvIndex) 
    {
        document.getElementById("pbGvRowIndex").value = gvIndex;
        submitAction("SHOW_SERIAL", "updtPnlPostBack");
    }


    function updtBatchNo(batchPattern, dateType, expCtrlID, manuCtrlID, batchCtrlID) {
        //alert(batchPattern + ' - ' + dateType + ' - ' + dateValue + ' - ' + batchCtrlID);

        //document.getElementById(manuCtrlID).value

        //document.getElementById(batchCtrlID).value

        //alert(document.getElementById(manuCtrlID).value.replace(/\//gi, ""));
        /*
        alert(("abcde").substring(0, 4));

        alert(batchCtrlID);


        var batchCtrl = $find(batchCtrlID);

        alert(batchCtrl.get_textBoxControl().value);


        batchCtrl.get_textBoxControl().value = "testing";

        alert(batchCtrl.get_textBoxControl().value);
        */
        //alert(document.getElementById(batchCtrlID).value);

        var batchCtrl = $find(batchCtrlID);


//        var ComboBoxCtrl = $find('combo1');
//         var item = document.createElement('LI');
//        item.innerHTML = batchCtrl.get_textBoxControl().value;
//         batchCtrl._optionListControl.appendChild(item);



        if (batchCtrl.get_textBoxControl().value == "" || batchCtrl.get_textBoxControl().value == "N/A" || batchCtrl.get_textBoxControl().value.substring(0, 4) == "@B#_")
        {
            if (batchPattern == "MANU_DATE")
            {
                if (document.getElementById(manuCtrlID).value == "")
                    batchCtrl.get_textBoxControl().value = "";
                else
                    batchCtrl.get_textBoxControl().value = "@B#_M_" + document.getElementById(manuCtrlID).value.replace(/\//gi, "");
            }
            else if (batchPattern == "EXP_DATE")
            {
                if (document.getElementById(expCtrlID).value == "")
                    batchCtrl.get_textBoxControl().value = "";
                else
                    batchCtrl.get_textBoxControl().value = "@B#_E_" + document.getElementById(expCtrlID).value.replace(/\//gi, "");
            }
            else if (batchPattern == "EXP_MANU")
            {
                if (document.getElementById(expCtrlID).value == "" && document.getElementById(manuCtrlID).value == "")
                    batchCtrl.get_textBoxControl().value = "";
                else if (document.getElementById(expCtrlID).value != "")
                    batchCtrl.get_textBoxControl().value = "@B#_E_" + document.getElementById(expCtrlID).value.replace(/\//gi, "");
                else if (document.getElementById(manuCtrlID).value != "")
                    batchCtrl.get_textBoxControl().value = "@B#_M_" + document.getElementById(manuCtrlID).value.replace(/\//gi, "");
            }
        }
    }

   function OpenNotes() {
        
    removeAllElementFromForm(document.hiddenForm);

    window.open("", "notes", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

    setInterfaceDataToForm(document.hiddenForm, "DOC_TYPE", "IB_GR");
    setInterfaceDataToForm(document.hiddenForm, "DOC_NO", document.getElementById("GR_CODE_HF").value);
    setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
   
    document.hiddenForm.action = "../../ALERT/NotesMain.aspx";
    document.hiddenForm.target = "notes";
    document.hiddenForm.submit();
}

    function getLoad() {
        var load_modalPopup = $find('load_ModalPopupExtender');
        load_modalPopup.show();
    }

$(document).keypress(
    function(event){
     if (event.which == '13') {
        event.preventDefault();
      }


});

</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0" onload="javascript:DisableDeleteButton();">
    <br />
    <form  id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div id="div1">
         <table border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
         <tr>
            <td>
        <table border="0" cellspacing="1" cellpadding="1" align="Left" style="width: 100%">
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
                    <td class="menuTD" align="left" style="vertical-align:top">
                        <asp:Button ID="selectROBtn" runat="server" Text="Select PO" CssClass="all_button" />
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" />                        
                        <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                            <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_GR'"
                            class="all_button" />
                    </td>
                    <td class="menuTD" align="right">
                        <input runat="server" id="btnRefresh" type="button" value="Refresh" class="all_button" onclick="Javascript:reloadPage('');" />
                         <asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" />    
                        <!--<input id="btnPickList" type="button" <%if Session("gLang") = "E" Then %>value="Put Away"
                        <% Elseif Session("gLang") = "C" Then %>value="Put Away" <% End If%> onclick="Javascript:putAwayMain();"
                        class="all_button" />-->
                        <asp:Button ID="btnInsp" runat="server" Text="Inspection" CssClass="all_button" OnClientClick="inspMain();return false;" />
                        <asp:Button ID="btnPutAway" runat="server" Text="Put Away" CssClass="all_button" />
                        <%If Session("pagemode") <> "N" Then%>             
                        <input id="btnPAList" type="button" <%if Session("gLang") = "E" Then %>value="Print Put Away list"
                        <% Elseif Session("gLang") = "C" Then %>value="印刷上架表" <% End If%> onclick="Javascript:OpenPAReport();"
                        class="all_button" />
                        <input id="btnStockRep" type="button" <%if Session("gLang") = "E" Then %>value="Stock Receipt"
                        <% Elseif Session("gLang") = "C" Then %>value="收貨表" <% End If%> onclick="Javascript:OpenGRReport();"
                        class="all_button" />
                        <%-- <input id="Button1" type="button" <%if Session("gLang") = "E" Then %>value="Cargo Receipt"
                        <% Elseif Session("gLang") = "C" Then %>value="Cargo Receipt" <% End If%> onclick="Javascript:OpenCargoRCP();"
                        class="all_button" />--%>                       
                        <input id="btnPalletLbl" type="button" <%if Session("gLang") = "E" Then %>value="Print Pallet Labels"
                        <% Elseif Session("gLang") = "C" Then %>value="托盤標籤" <% End If%> onclick="Javascript:OpenPalletLbls();"
                        class="all_button" />
                        <input id="btnCartonLbl" type="button" <%if Session("gLang") = "E" Then %>value="Carton Labels"
                        <% Elseif Session("gLang") = "C" Then %>value="紙箱標籤" <% End If%> onclick="Javascript:OpenCartonLbls();"
                        class="all_button" <% If editMode.Value = "N" Then%>disabled<% End If%> />
                        <input id="btnitmLabel" type="button" <%if Session("gLang") = "E" Then %>value="Item Labels"
                        <% Elseif Session("gLang") = "C" Then %>value="貨品標籤" <% End If%> onclick="Javascript:OpenItemLbls();"
                        class="all_button" <% If editMode.Value = "N" Then%>disabled<% End If%> />
                        <br />
                        <asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" />
                        <asp:Button ID="btnInTransit" Text="In Transit" CssClass="all_button" Visible="false" runat="server" />
                        <asp:Button ID="btnRelease" Text="Release" CssClass="all_button" Visible="false" runat="server" />
                        <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" />
                        <asp:Button ID="btnUnPost" Text="Un-Post" CssClass="all_button" runat="server" />                                           
                        <% End If%>
                         
                    </td>
                </tr>
                </table>
                    
                </td>
            </tr>
            <tr>
                <td colspan="6" class="TITLE">
                    <asp:UpdatePanel runat="server" ID="updtPnlLinkBar" UpdateMode="Conditional" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:Table ID="linkBar" runat ="server" align="center" width="757"></asp:Table>
                    </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
            </tr>               
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_CODE" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:Label ID="GR_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="GR_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" MaxLength="20" AutoPostBack="true" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_EDI_PO_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="GR_EDI_PO_NO" runat="server" MaxLength="80" Width="150px"></asp:TextBox>
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="GR_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="GR_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_RCV_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="GR_RCV_BY" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_WH_CODE" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="GR_WH_CODE" CssClass="READONLY" runat="server" Width="150px" />
                </td>                
            </tr>
            <tr runat="server" visible="false">
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lblBarcode" runat="server" Text="Barcode:" /></font>
                </td>
                <td id="Td1" runat="server" nowrap colspan="5">
                    <font size="2">
                        <table cellpadding="0" cellspacing="0">
                            <tr>
                                <td>
                                    <asp:TextBox ID="txtBarCode" runat="server" Width="250px"></asp:TextBox>            
                                </td>
                                <td>
                                    &nbsp;&nbsp;
                                    <asp:Label ID="txtOutput" runat="server" />
                                </td>
                            </tr>
                        </table>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_PRJ_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="PRJ_CODE" runat="server" MaxLength="20"></asp:DropDownList>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                   <asp:Label ID="lbl_GR_TRANS_TYPE" runat="server" Text="Nature of Transaction"  />:
                </td>
                <td><font size="2">
                    <asp:DropDownList runat="server" ID="GR_TRANS_TYPE" />
                </font></td>
                  <td class="LabelTD" nowrap>
                   <asp:Label ID="lbl_gr_cat" runat="server" Text="Category"  />:
                </td>
                <td><font size="2">
                    <asp:DropDownList runat="server" ID="gr_cat" />
                </font></td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_DOC_TYPE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="txt_GR_DOC_TYPE" runat="server"></asp:Label>
                        <asp:HiddenField ID="GR_DOC_TYPE" runat="server"></asp:HiddenField>                        
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_DOC_NO" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="txt_GR_DOC_NO" runat="server"></asp:Label>
                        <asp:HiddenField ID="GR_DOC_NO" runat="server"></asp:HiddenField>
                    </font>
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_PO_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="PO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                      <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="PO_DATE" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_VND_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                    <%--<asp:Label ID="dsp_VND_CODE" runat="server" />--%>
                        <asp:TextBox ID="dsp_VND_CODE" runat="server" MaxLength="30"></asp:TextBox>
                    <asp:HiddenField ID="VND_CODE" runat="server" value="" />
                    </font>
                    <img visible="false" onclick="javascript:VendorLookUp();" onmouseout="MM_swapImgRestore()" onmousedown="MM_swapImage('Image_Vnd_LookUp','','../../images/btn_search_over.gif',1)" style="cursor:hand" align="absmiddle" src="../../images/btn_search.gif" name="Image_Vnd_LookUp" border="0"  runat="server">                    
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_VND_NAME" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="dsp_VND_NAME" runat="server" />
                        <asp:HiddenField ID="VND_NAME" runat="server" value="" />
                    </font>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_VND_DNREF" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="GR_VND_DNREF" runat="server" MaxLength="30"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_DESTINATION" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="GR_DESTINATION" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_TOT_PALLET" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="GR_TOT_PALLET" runat="server" MaxLength="12" style="text-align:right"></asp:TextBox>
                    </font>
                </td>
            </tr> 
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_HAWB" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="GR_HAWB" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_INV_NO" runat="server" /></font>
                </td>
                <td nowrap colspan="3">
                    <font size="2">
                        <asp:TextBox ID="GR_INV_NO" runat="server" MaxLength="50"></asp:TextBox>
                    </font>
                </td>
            </tr>                       
            
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_TRK_NO" runat="server" >TRACK No</asp:Label>
                    </font>
                </td>
                <td colspan="5">
                    <font size="2">
                        <%--<asp:Label ID="txt_TRK_NO" runat="server"></asp:Label>--%>
                        <asp:TextBox ID="txt_TRK_NO" runat="server" MaxLength="400"></asp:TextBox>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_GR_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="GR_REM" runat="server" Height="117px" Width="384px" TextMode="MultiLine"></asp:TextBox>
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
         </td>
         </tr>
         </table>
        <table border="0" cellspacing="1" cellpadding="1" align="Left" style="width:100%">
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
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible="false"  />
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" Visible="false" />
                    <asp:Button ID="btnItemDelete" runat="server" Text="Delete Items" CssClass="all_button" Visible="true" />
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="6">
                    <asp:UpdatePanel runat="server" ID="updtPnlItemGV" UpdateMode="Conditional" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                            Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                            CellPadding="3" CaptionAlign="Top" HorizontalAlign="Left" ShowFooter="true">
                            <Columns>
                                <asp:TemplateField ControlStyle-Width="30px">
                                    <ItemTemplate>
                                        <asp:CheckBox ID="cb_select" runat="server" />
                                    </ItemTemplate>
                                    <ControlStyle Width="30px"></ControlStyle>
                                    <HeaderStyle Width="30px" />
                                </asp:TemplateField>
                                <asp:TemplateField ControlStyle-Width="30px">
                                    <ItemTemplate>
                                        <asp:Button ID="btnSplit" name="btnSplit" runat="server" Height="22px" Font-Size="11px"
                                            CommandName="SplitItem" Text="S" CssClass="all_button" Font-Bold="false" />
                                    </ItemTemplate>
                                    <ControlStyle Width="30px"></ControlStyle>
                                    <HeaderStyle Width="30px" />
                                </asp:TemplateField>
                                <asp:TemplateField ControlStyle-Width="24px">
                                    <ItemTemplate>
                                        <asp:Button ID="btnHighlight" name="btnHighlight" runat="server" Height="22px" Font-Size="11px" Text=">" CssClass="all_button" Font-Bold="false" />
                                    </ItemTemplate>
                                    <ControlStyle Width="24px"></ControlStyle>
                                    <HeaderStyle Width="24px" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Seq No.">
                                    <ControlStyle Width="30px" />
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_disp_seq" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" onkeypress="return maskNumOnly(event);" />
                                        <asp:HiddenField ID="grd_seq" runat="server" />
                                        <asp:HiddenField ID="grd_ref_seq" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Vendor Code">
                                    <ControlStyle Width="120px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_grd_vnd_code" runat="server" Font-Size="11px" Width="120px"></asp:Label>
                                        <asp:HiddenField ID="grd_vnd_code" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Item Code">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <table border="0" cellpadding ="0" cellspacing ="0" width="100%" runat="server" id="ser_table">
                                        <tr>
                                            <td align="left">
                                                 <asp:Label ID="grd_itm_code" width="120" MaxLength="20" runat="server" Font-Size="11px" />
                                            </td>
                                            <td align="right">
                                                 <asp:ImageButton ID="btnItemSerial" runat="server" ImageUrl="../../images/s_dis_icon.gif" CommandName="SHOW_SERIAL" style="border-width:0px;cursor:hand" align="absmiddle" />
                                            </td>
                                        </tr>
                                        </table>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="false" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Stock No.">
                                    <ControlStyle Width="120px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="itm_sku_no" width="100" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pack Key">
                                    <ControlStyle Width="40px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="grd_pack_key" width="100" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Item Name">
                                    <ControlStyle Width="120px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:Label ID="grd_itm_name" width="100" MaxLength="100" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pack Type">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:DropDownList ID="grd_pack" runat="server" Font-Size="11px">
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pack No.">
                                    <ControlStyle Width="80px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_pack_no" width="70" MaxLength="20" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="PO Qty">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_po_qty" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);" ReadOnly=true />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="RCV Qty" AccessibleHeaderText="grd_rcv_qty">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_rcv_qty" CssClass="REQUIRED" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_os_qty" MaxLength="0" runat="server" Font-Size="0px" style="height:0px; width: 0px;BORDER-TOP-WIDTH: 0px;BORDER-LEFT-WIDTH: 0px;BORDER-BOTTOM-WIDTH: 0px;BORDER-RIGHT-WIDTH: 0px"/>
                                    </ItemTemplate>
                                    <FooterTemplate>
                                       Total:
                                    </FooterTemplate>
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Pallet No.">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_pallet_no" width="70" MaxLength="20" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="No. of Carton" AccessibleHeaderText="grd_no_of_carton">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_no_of_carton" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);" />
                                    </ItemTemplate>
                                    <FooterTemplate>
                                       <asp:Label ID="grd_total_carton" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:Label>
                                    </FooterTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Qty Per Carton">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_pcs_per_carton" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskKey(event);" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Carton No.">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_carton_no" width="80" MaxLength="20" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Batch No.">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                     <%--<asp:ComboBox ID="grd_batch_no" runat="server" Font-Size="11px"  
                                                      MaxLength="20" AutoCompleteMode="SuggestAppend" DropDownStyle="DropDown" ItemInsertLocation="OrdinalText" CaseSensitive="false">
                                        </asp:ComboBox>  --%>              
                                        <asp:TextBox ID="grd_batch_no" width="80" MaxLength="20" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>                           
                                <asp:TemplateField HeaderText="Reference">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_ref_no" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                               <%-- <asp:TemplateField HeaderText="Brand">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_brand" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Series">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_series" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Model">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_model" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>--%>
                                <asp:TemplateField HeaderText="UOM">
                                    <ControlStyle Width="100px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:DropDownList ID="grd_uom" runat="server" Font-Size="11px" Enabled="false" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Size (LxWxH)" AccessibleHeaderText="grd_size">
                                    <ControlStyle Width="40px" />
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" Wrap="false" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_length" width="30" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right"></asp:TextBox>&nbsp;
                                        <asp:TextBox ID="grd_width" width="30" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right"></asp:TextBox>&nbsp;
                                        <asp:TextBox ID="grd_height" width="30" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Weight(kg)" AccessibleHeaderText="grd_kg_c">
                                    <ControlStyle Width="40px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_kg" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Sub-Total(kg)" AccessibleHeaderText="grd_kg">
                                    <ControlStyle Width="40px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                    <ItemTemplate>
                                        <asp:Label ID="sub_grd_kg" width="40px" Font-Size="11px" MaxLength="12" runat="server" style="text-align:right" />
                                    </ItemTemplate>
                                    <FooterTemplate>
                                       <asp:Label ID="grd_total_kg" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:Label>
                                    </FooterTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Volume(cbm)" AccessibleHeaderText="grd_cbm">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_cbm" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                    <FooterTemplate>
                                       <asp:Label ID="grd_total_cbm" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:Label>
                                    </FooterTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Destination">
                                    <ControlStyle Width="100px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_destination" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Stock Type">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_stocktype" width="70" MaxLength="20" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                               <%--  <asp:TemplateField HeaderText="Stock Category">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" ID="grd_cat" Font-Size="11px"/>
                                </ItemTemplate>                          
                            </asp:TemplateField>--%>
                                <asp:TemplateField HeaderText="Number per UOM">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_pcs_per_uom" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Total Number" AccessibleHeaderText="grd_tot_pcs">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <FooterStyle HorizontalAlign="Right" Font-Bold="true" Font-Size="11px" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_tot_pcs" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                    <FooterTemplate>
                                       <asp:Label ID="grd_sub_total_pcs" runat="server" Font-Size="11px" Width="40px" style="text-align:right"></asp:Label>
                                    </FooterTemplate>    
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Expiry Date">
                                    <ControlStyle />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle Wrap="false" />
                                    <ItemTemplate>
                                         <asp:TextBox ID="GRD_EXPIRY_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                         <asp:ImageButton ID="btnCal01" runat="server" ImageUrl="../../images/calendar1.gif" 
                                                ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                         <asp:CalendarExtender ID="calEXP" runat="server" CssClass="ajax_calendar" 
                                                TargetControlID="GRD_EXPIRY_DATE" PopupButtonID="btnCal01" Format="dd/MM/yyyy" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Manufactory Date">
                                    <ControlStyle />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle Wrap="false" />
                                    <ItemTemplate>
                                         <asp:TextBox ID="GRD_MANU_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskDate(event);"></asp:TextBox>
                                         <asp:ImageButton ID="btnCal02" runat="server" ImageUrl="../../images/calendar1.gif" 
                                                ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                         <asp:CalendarExtender ID="calMANU" runat="server" CssClass="ajax_calendar" 
                                                TargetControlID="GRD_MANU_DATE" PopupButtonID="btnCal02" Format="dd/MM/yyyy" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Rejected Qty.">
                                    <HeaderStyle  />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:TextBox ID="GRD_REJ_QTY" width="50" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Reject Reason">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:DropDownList runat="server" ID="GRD_REJ_REASON"  Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Rej. Remark">
                                    <HeaderStyle  />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:TextBox ID="GRD_REJ_RMKS" width="80" MaxLength="500" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Insp. Req.">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_GRD_INSP_REQ" Font-Size="11px" runat="server" />
                                        <asp:HiddenField ID="GRD_INSP_REQ" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Status">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:DropDownList runat="server" ID="GRD_STATUS" Font-Size="11px" Visible="false" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="UOM2">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:DropDownList runat="server" ID="GRD_UOM2"  Font-Size="11px" Enabled="false" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Qty2">
                                    <ControlStyle Width="50px" />
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="GRD_QTY2" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <%--<asp:TemplateField HeaderText="On Behalf">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:Label ID="dsp_GRD_ON_BEHALF" Font-Size="11px" runat="server" />
                                        <asp:HiddenField ID="GRD_ON_BEHALF" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>--%>
                                <asp:TemplateField HeaderText="Warehouse">
                                    <HeaderStyle />
                                    <ItemStyle />
                                    <ItemTemplate>
                                        <asp:Label ID="GRD_WH" Font-Size="11px" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                               <%-- <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnCopySize" name="btnCopySize" runat="server" Height="22px" Font-Size="11px"
                                            CommandName="Update" Text="C" CssClass="all_button" Font-Bold="false" Width="30px" />     
                                    </ItemTemplate>
                                    <ControlStyle Width="30px"></ControlStyle>
                                    <HeaderStyle Width="30px" />
                                </asp:TemplateField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                         <asp:Button ID="btnPaste" name="" runat="server" Height="22px" 
                                            Font-Size="11px" Text="P" CssClass="all_button" Width="30px"
                                            Font-Bold="false"    />
                                    </ItemTemplate>
                                    <ControlStyle Width="30px"></ControlStyle>
                                    <HeaderStyle Width="30px" />
                                </asp:TemplateField>--%>
                            
                            </Columns>
                             <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                        </asp:GridView>
                    </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" />
                    <input id="btnBack" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../cms_search.aspx?menu_code=IB_GR'"
                        class="all_button" />
                </td>
            </tr>
        </table>
    </div>


    <asp:Panel ID="pnlDetailSerial" runat="server" CssClass="modalPopup" style="display: none">
    <table width="500px" cellpadding="0" cellspacing="0" border="0">
    <tr style="height:25px">
        <td class="TITLE">
            <asp:Panel runat="Server" ID="pnlDetailSerialDrag" Style="cursor: move;">
            <font size="2">Serial</font>
            </asp:Panel>
        </td>
    </tr>
        <tr>
        <td>
            <asp:UpdatePanel runat="server" ID="updtPnl_DetailSerialItem" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
            <table width="100%" cellpadding="0" cellspacing="1" border="0">
            <tr style="height:22px">
                <td class="LabelTD" width="20%">Item Code:</td>
                <td width="30%">
                    <asp:Label ID="DS_ITEM_CODE" runat="server" />
                </td>
                <td class="LabelTD" width="20%">Pack Key:</td>
                <td>
                    <asp:Label ID="DS_PACK_KEY" runat="server" />
                </td>
            </tr>
            <tr style="height:22px">
                <td class="LabelTD">Stock No.:</td>
                <td colspan="3">
                    <asp:Label ID="DS_SKU_NO" runat="server" />
                </td>
            </tr>
            <tr style="height:22px">
                <td class="LabelTD">Batch No.:</td>
                <td>
                    <asp:Label ID="DS_BATCH_NO" runat="server" />
                </td>
                <td class="LabelTD">Pallet No.:</td>
                <td>
                    <asp:Label ID="DS_PALLET_NO" runat="server" />
                </td>
            </tr>
            <tr style="height:22px">
                <td class="LabelTD">Item Name:</td>
                <td colspan="3">
                    <asp:Label ID="DS_ITEM_NAME" runat="server" />
                    <asp:HiddenField ID="DS_WH" runat="server" />
                </td>
            </tr>
            </table>
            </ContentTemplate>
            </asp:UpdatePanel>
        </td>
    </tr>
    <tr>
        <td class="TITLE" align="left">
            <asp:Button runat="server" ID="btnAddDetailSerial" Text="Add Serial No." CssClass="all_button" />&nbsp;
        </td>            
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td>
            <asp:UpdatePanel runat="server" ID="updtPnl_DetailSerialList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>
                <asp:GridView ID="gvDetailSerialList" runat="server" Width="100%" Font-Names="Arial"
                    Font-Overline="False" Font-Size="10px" 
                    AutoGenerateColumns="False" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" ShowHeader="True"
                    CellPadding="2" CaptionAlign="Top" GridLines="None" CellSpacing="1" ShowHeaderWhenEmpty="true" 
                    HorizontalAlign="Left">
                    <RowStyle CssClass="GV" />
                    <HeaderStyle CssClass="DtlLabel" Font-Bold="False"/>
                    <Columns>
                    <asp:TemplateField HeaderText="Serial No.">
                        <ControlStyle Width="100px" />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <asp:TextBox ID="GRS_SERIAL_NO" width="100px" MaxLength="80" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Location">
                        <ControlStyle />
                        <HeaderStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                        <asp:DropDownList ID="GRS_LOC" runat="server" Font-Size="11px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" Width="100px" />
                        <ItemTemplate>
                            <asp:Button ID="btnDelDtl" Text="Delete" CssClass="all_button" CommandName="DELETE" runat="server" />
                            <asp:HiddenField ID="GRS_SEQ" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateField>
                    </Columns>
                    </asp:GridView>
            </ContentTemplate>
            </asp:UpdatePanel>    
            <asp:UpdateProgress runat="server" ID="prog1"  DisplayAfter="0">
                <ProgressTemplate>
                    <div style="width:100%; text-align:center;"> <asp:Image runat="server" ID="loadbar" ImageUrl="~/images/loadbar.gif" /> </div>
                </ProgressTemplate>
            </asp:UpdateProgress>        
        </td>
    </tr>
    <tr>
        <td class="TITLE" align="left">
            <asp:Button runat="server" ID="btnDetailSerialOK" Text="OK" CssClass="all_button" />&nbsp;
            <asp:Button runat="server" ID="btnCloseDetailSerial" Text="Close" CssClass="all_button" />
        </td>
    </tr>
    </table>
    </asp:Panel>

    <asp:ModalPopupExtender ID="DetailSerial_ModalPopupExtender" runat="server"
        DynamicServicePath="" 
        Enabled="True" 
        TargetControlID="detailSerial_dummy" 
        PopupControlID="pnlDetailSerial"
        BackgroundCssClass="modalBackground"
        DropShadow="true"         
        CancelControlID="btnCloseDetailSerial"
        Y="50"
        PopupDragHandleControlID="pnlDetailSerialDrag" 
        RepositionMode="None" 
        BehaviorID="detailSerial_behavior"
        />

    <asp:HiddenField ID="detailSerial_dummy" runat="server" /> 


    <!-- Update panel for asyn postback -->
    <asp:UpdatePanel runat="server" ID="updtPnlPostBack" UpdateMode="Conditional">
    <ContentTemplate />
    </asp:UpdatePanel>

    <!-- Update panel that decided for validation alert -->
    <asp:UpdatePanel runat="server" ID="updtPnlAlert">
    <ContentTemplate />
    </asp:UpdatePanel>


    <!-- Update panel that decided for key fields and moduleAction -->
    <asp:UpdatePanel runat="server" ID="updtPnlKey">
    <ContentTemplate>
        <asp:HiddenField ID="IMP_CODE" runat="server" />
        <asp:HiddenField ID="editMode" runat="server" />
        <asp:HiddenField ID="C_L" runat="server" />
        <asp:HiddenField ID="C_W" runat="server" />
        <asp:HiddenField ID="C_H" runat="server" />
        <asp:HiddenField ID="C_KG" runat="server" />
        <asp:HiddenField ID="sel_row_idx" runat="server" />
        <asp:HiddenField ID="FocusBarYN" runat="server" Value="Y" />

        <asp:HiddenField ID="GR_CODE_HF" runat="server" />

        <asp:HiddenField ID="GR_TRACK_NO" runat="server" />
        <asp:HiddenField ID="GR_REF_NO" runat="server" />

        <asp:HiddenField ID="moduleAction" runat ="server" />
        <asp:HiddenField ID="ROList" runat ="server" />
        <asp:HiddenField ID="ROSeqList" runat ="server" />
        <asp:HiddenField ID="itemList" runat ="server" />
        <asp:HiddenField ID="packKeyList" runat ="server" />            
        <asp:HiddenField ID="dummy" runat ="server" />            
        <asp:HiddenField ID="qtyList" runat ="server" />

        <asp:HiddenField ID="pbGvRowIndex" runat="server" />
        
    </ContentTemplate>
    </asp:UpdatePanel>


    <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
                ScrollBars="None">
                <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
                    height: 80px">
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

    </form>
    <form name="hiddenForm" id="hiddenForm" method="post" />   
    
    <%-- <asp:TemplateField HeaderText="Brand">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_brand" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Series">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_series" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Model">
                                    <ControlStyle Width="70px" />
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="grd_model" width="70" MaxLength="50" runat="server" Font-Size="11px" />
                                    </ItemTemplate>
                                </asp:TemplateField>--%>
</body>
</html>
