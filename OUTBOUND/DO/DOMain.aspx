<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DOMain.aspx.vb" Inherits="DOMain" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <title>Stock Return</title>
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

        if (document.myform.editMode.value != "V") {
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

    function LocLookUp(lb_id, hd_id) {
        if (document.myform.editMode.value != "V") {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "locLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=700,left=5,top=15,resizable=yes");

            setInterfaceDataToForm(document.hiddenForm, "wh", document.myform.RT_WH.value);
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pItemList", lb_id + "|L, " + hd_id);
            document.hiddenForm.action = "../../LOOKUP/locLookup.aspx";
            document.hiddenForm.target = "locLookUp";
            document.hiddenForm.submit();
        }
    }

    function packingMain() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "packingMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1280,height=400,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "do_status", document.getElementById('DO_STATUS').innerHTML);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_1", document.myform.DO_PACK_LABEL_1.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_2", document.myform.DO_PACK_LABEL_2.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_3", document.myform.DO_PACK_LABEL_3.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_QTY_1", document.myform.DO_PACK_LABEL_QTY_1.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_QTY_2", document.myform.DO_PACK_LABEL_QTY_2.value);
        setInterfaceDataToForm(document.hiddenForm, "DO_PACK_LABEL_QTY_3", document.myform.DO_PACK_LABEL_QTY_3.value);
        document.hiddenForm.action = "PackingList.aspx";
        document.hiddenForm.target = "packingMain";
        document.hiddenForm.submit();
    }

    function loadInitJS() {
        if (document.getElementById('INIT_JS').value == "POST_CHK")
            postChk();
    }

    function pickListMain() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "pickListMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1280,height=600,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "co_code", document.myform.DO_CO_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        setInterfaceDataToForm(document.hiddenForm, "do_status", document.getElementById('DO_STATUS').innerHTML);
        document.hiddenForm.action = "PickList.aspx";
        document.hiddenForm.target = "pickListMain";
        document.hiddenForm.submit();
    }

    function postChkReload() {
        removeAllElementFromForm(document.hiddenForm);
        setInterfaceDataToForm(document.hiddenForm, 'DO_CODE', document.getElementById('DO_CODE_HF').value);

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("STORER_CODE").value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "READLOAD_JS", "POST_CHK");

        document.hiddenForm.action = './DOMain.aspx';
        document.hiddenForm.target = '_self';
        document.hiddenForm.submit();
    }

    function postChk() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "pickListMain", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=600,height=400,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "co_code", document.myform.DO_CO_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        setInterfaceDataToForm(document.hiddenForm, "do_status", document.getElementById('DO_STATUS').innerHTML);
        document.hiddenForm.action = "PostCheck.aspx";
        document.hiddenForm.target = "pickListMain";
        document.hiddenForm.submit();
    }

    function openLabel() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "labelWin", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "do_status", document.getElementById('DO_STATUS').innerHTML);
        document.hiddenForm.action = "LabelList.aspx";
        document.hiddenForm.target = "labelWin";
        document.hiddenForm.submit();
    }

    function checkAvail() {
        if (confirm("The checking will check with saved data only. Please make sure you have saved the order.")) {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "availList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=800,height=300,left=5,top=15");

            if (document.myform.editMode.value != "V") {
                setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
            } else {
                setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
            }

            setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
            setInterfaceDataToForm(document.hiddenForm, "co_code", document.myform.DO_CO_CODE.value);
            setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
            document.hiddenForm.action = "Availability.aspx";
            document.hiddenForm.target = "availList";
            document.hiddenForm.submit();
        }
    }

    function COLookUp(txt_GR_DOC_NO, GR_DOC_NO, STORER_CODE) {
        if (STORER_CODE == '') {
            alert('Please select the Storer first!');
            return false;
        }

        if (document.myform.editMode.value != "V") {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "COLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=750,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "menu_code", "LOOKUP_CO");
            setInterfaceDataToForm(document.hiddenForm, "pForm", "myform");
            setInterfaceDataToForm(document.hiddenForm, "pFunc", "selectedCO()");
            setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
            //setInterfaceDataToForm(document.hiddenForm, "pItemList", GR_DOC_NO + "|1");
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "COList|1, COSeqList|2, qtyList|CC_ITEM_QTY#txt_item_qty");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "COLookUp";
            document.hiddenForm.submit();
        }
    }

    function PrintDDN() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintDNN", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=800,height=450,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        document.hiddenForm.action = "D_DELI_NOTE/DriverDeliNote.aspx";
        document.hiddenForm.target = "doPrintDNN";
        document.hiddenForm.submit();
    }

    function PrintPackList() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintPackList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=840,height=680,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        document.hiddenForm.action = "PACK_LIST/PackList.aspx";
        document.hiddenForm.target = "doPrintPackList";
        document.hiddenForm.submit();

    }


    function PrintPickList() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintPickList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=1150,height=576,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        document.hiddenForm.action = "PICK_LIST/picklist_print.aspx";
        document.hiddenForm.target = "doPrintPickList";
        document.hiddenForm.submit();

    }

    function PrintDON() {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPrintDON", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=1150,height=700,left=5,top=15");
        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById('DO_CODE_HF').value);
        document.hiddenForm.action = "DELI_ORDER/DeliOrderNote.aspx";
        document.hiddenForm.target = "doPrintDON";
        document.hiddenForm.submit();
    }

    function selectedCO() {
        document.myform.moduleAction.value = "SELECTCO";
        document.myform.submit();
    }

    function reloadPackList() {
        document.myform.moduleAction.value = "RELOADPACK";
        document.myform.submit();
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
            setInterfaceDataToForm(document.hiddenForm, "pItemList", "itemList|1, packKeyList|2");
            document.hiddenForm.action = "../../cms_search.aspx";
            document.hiddenForm.target = "ItemLookUp";
            document.hiddenForm.submit();
        }
    }

    function selectedItem() {
        document.myform.moduleAction.value = "SELECTIM";
        document.myform.submit();
    }

    function checkSB(STORER_CODE) {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "sbLookUp", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=1200,height=800,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "menu_code", "INQ_001");
        setInterfaceDataToForm(document.hiddenForm, "sc", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "ctemp", true);
        setInterfaceDataToForm(document.hiddenForm, "screadonly", true);

        document.hiddenForm.action = "../../cms_search.aspx";
        document.hiddenForm.target = "sbLookUp";
        document.hiddenForm.submit();
    }

    function saveok() {
        /*
        document.myform.moduleAction.value = "SAVEOK";
        document.myform.submit();
        */
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

    function reloadPage(alertMsg) {
        if (alertMsg)
            alert(alertMsg);

        removeAllElementFromForm(document.hiddenForm);
        setInterfaceDataToForm(document.hiddenForm, 'DO_CODE', document.getElementById('DO_CODE_HF').value);

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", document.getElementById("STORER_CODE").value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", "<%=ViewState(Page.ClientID & "_PAGE_STORER_CODE")%>");
        }

        document.hiddenForm.action = './DOMain.aspx';
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

    function OpenItemLbls() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doItemLbls", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById("DO_CODE").innerHTML);
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);

        document.hiddenForm.action = "DO_LABELS/item_label_print.aspx";
        document.hiddenForm.target = "doItemLbls";
        document.hiddenForm.submit();
    }

    //Short Ship Report
    function OpenShortShipReport() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doShortShip", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById("DO_CODE").innerHTML);
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);

        document.hiddenForm.action = "SHORT_SHIP/short_ship.aspx";
        document.hiddenForm.target = "doShortShip";
        document.hiddenForm.submit();
    }


    //Picking List Report
    function OpenPickingListReport() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "doPickingList", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "do_code", document.getElementById("DO_CODE").innerHTML);
        setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);

        document.hiddenForm.action = "PICKING_LIST/picking_list.aspx";
        document.hiddenForm.target = "doPickingList";
        document.hiddenForm.submit();
    }

    function openPackLabel(do_code, dov_seq) {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "PLabel", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,resizable=yes,personalbar=0,status=1,width=700,height=850,left=5,top=15");

        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.myform.STORER_CODE.value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        setInterfaceDataToForm(document.hiddenForm, "imp_code", document.myform.IMP_CODE.value);
        setInterfaceDataToForm(document.hiddenForm, "do_code", do_code);
        setInterfaceDataToForm(document.hiddenForm, "dov_seq", dov_seq);

        document.hiddenForm.action = "PACK_LABEL/PackLabel.aspx";
        document.hiddenForm.target = "PLabel";
        document.hiddenForm.submit();
    }


    function getLoad() {
        var load_modalPopup = $find('load_ModalPopupExtender');
        load_modalPopup.show();
    }

    function postOrderOK() {
        //alert('Post DO gogogo!');
        submitAction("POSTOK", "updtPnlPostBack");
    }


    function OpenNotes() {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "notes", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "DOC_TYPE", "OB_DO");
        if (document.myform.editMode.value != "V") {
            setInterfaceDataToForm(document.hiddenForm, "DOC_NO", document.getElementById("DO_CODE").innerHTML);
            setInterfaceDataToForm(document.hiddenForm, "storer_code", document.getElementById("STORER_CODE").value);
        } else {
            setInterfaceDataToForm(document.hiddenForm, "DOC_NO", document.getElementById("DO_CODE").innerHTML);
            setInterfaceDataToForm(document.hiddenForm, "storer_code", "<%=Session("STORER_CODE")%>");
        }

        document.hiddenForm.action = "../../ALERT/NotesMain.aspx";
        document.hiddenForm.target = "notes";
        document.hiddenForm.submit();
    }

    function OpenTrack(sto_code, ftrack_no) {

        removeAllElementFromForm(document.hiddenForm);

        window.open("", "track", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,status=1,width=900,height=600,left=5,top=15");

        setInterfaceDataToForm(document.hiddenForm, "ftrack_no", ftrack_no);
        setInterfaceDataToForm(document.hiddenForm, "storer_code", sto_code);

        document.hiddenForm.action = "../../DELI_STATUS/DELI_STATUS.aspx";
        document.hiddenForm.target = "track";
        document.hiddenForm.submit();
    }



    /*
    $(document).keypress(
        function(event){
         if (event.which == '13') {
            event.preventDefault();
          }
    
    
    });
    
    */
</script>
</div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:DisableDeleteButton();loadInitJS();">
    <br />
    <form id="myform" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <%--<input type="hidden" name="moduleAction" value=""/>--%>

    <div id="div1">
        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 900px">
            <tr>
                <td colspan="6" class="TITLE">
                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                        <td class="TITLE" align="left">
                            <asp:Table ID="cmBar" runat ="server" border="0" cellspacing="0" cellpadding="0" />
                        </td>
                        <td class="TITLE" align="right">
                            <asp:Button runat="server" ID="btnNOTE" text="Notes" CssClass="all_button" OnClientClick="javascript:OpenNotes();return false;" Visible="false" />
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
                        <asp:Button ID="selectCOBtn" runat="server" Text="Select CO" CssClass="all_button" />
                        <asp:Button ID="saveStorerBtn" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                        <asp:Button ID="saveConDate" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" Visible="false"/>
                        
                        <input id="btnBack2" runat="server" type="button" value="Back" onclick="Javascript: window.location = '../../cms_search.aspx?menu_code=OB_DO'"
                            class="all_button" />
                        <input type="button" runat="server" id="BtnClose1" value="Close" onclick="Javascript: window.close();" class="all_button" />
                        
                    </td>
                    <td class="menuTD" align="right">
                        <input runat="server" id="btnRefresh" type="button" value="Refresh" class="all_button" onclick="Javascript: reloadPage();" />
                        <input id="btnPickList" type="button" <%if Session("gLang") = "E" Then %>value="Picking"
                            <% Elseif Session("gLang") = "C" Then %>value="執貨" <% End If%> onclick="Javascript:pickListMain();"
                            class="all_button" />
                          
                        <input id="btnPacking" type="button" <%if Session("gLang") = "E" Then %>value="Packing"
                            <% Elseif Session("gLang") = "C" Then %>value="封裝" <% End If%> onclick="Javascript: packingMain();"
                            class="all_button" />
                        <%If Session("pagemode") <> "N" Then%>
                         <asp:Button ID="btnPickingListReport" Text="Picking List Report" CssClass="all_button" runat="server" OnClientClick="javascript:OpenPickingListReport();return false;"/>
                           <asp:Button ID="btnSSReport" Text="Short Ship Report" CssClass="all_button" runat="server" OnClientClick="javascript:OpenShortShipReport();return false;"/>
                         <asp:Button ID="btnPrintPalletLabel" Text="Print Pallet Label" CssClass="all_button" runat="server" />
                            <asp:Button runat="server" ID="btnPrtLbl" Text="Print Item Label" CssClass="all_button" OnClientClick="javascript:OpenItemLbls();return false;" />
                            <input id="btnPrintDDn" type="button" <%if Session("gLang") = "E" Then %>value="Print DDN"
                                <% Elseif Session("gLang") = "C" Then %>value="打印司機送貨單" <% End If%> onclick="Javascript: PrintDDN();"
                                class="all_button" />
                            <input id="btnDeliNote" type="button" <%if Session("gLang") = "E" Then %>value="Delivery Note"
                                <% Elseif Session("gLang") = "C" Then %>value="送貨單" <% End If%> onclick="Javascript: PrintDON();"
                                class="all_button" />
                            <input id="Button1" type="button" <%if Session("gLang") = "E" Then %>value="Picking List"
                                <% Elseif Session("gLang") = "C" Then %>value="列印執貨表" <% End If%> onclick="Javascript:PrintPickList();"
                                class="all_button"  />
                            <input id="btnPackList" type="button" <%if Session("gLang") = "E" Then %>value="Packing List"
                                <% Elseif Session("gLang") = "C" Then %>value="列印封裝表" <% End If%> onclick="Javascript:PrintPackList();"
                                class="all_button"  />
                            <br />
                         <asp:Button ID="btnGenPickingList" Text="Generate Picked List" CssClass="all_button" runat="server" OnClientClick="return confirm(&quot;Generate Picked Recrods?&quot;);" />
                       <%-- <asp:Button ID="btnPicking" Text="Picking" CssClass="all_button" runat="server" />
                           <asp:Button ID="btnRelease" Text="Release" CssClass="all_button" runat="server" />--%>
                            <%If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then%>
                            <%ELSE %>
                            <input id="btnAvbChk" type="button" <%if Session("gLang") = "E" Then %>value="Availability"
                                <% Elseif Session("gLang") = "C" Then %>value="可用" <% End If%> onclick="Javascript:checkAvail();"
                                class="all_button" <%if DO_STATUS.Text = "POSTED" Or DO_STATUS.Text = "POSTED" Then%>disabled<% end if %> />
                            <% end if %>
                            <asp:Button ID="cSBBtn" Text="Check Stock Balance" CssClass="all_button" runat="server" />
                            <%--<input id="Button1" type="button" <%if Session("gLang") = "E" Then %>value="Label"
                                <% Elseif Session("gLang") = "C" Then %>value="標籤" <% End If%> onclick="Javascript:openLabel();"
                                class="all_button"  />--%>
                         
                            <asp:Button runat="server" id="btnPick" CssClass="all_button" Text="Pick" />
                            <asp:Button runat="server" id="btnUnPick" CssClass="all_button" Text="Unpick" />
                            <asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server" Visible="false" />
                            <asp:Button ID="btnPostChk" Text="Post" CssClass="all_button" runat="server" />
                            <asp:Button ID="btnUnPost" Text="Un-Post" CssClass="all_button" runat="server" />                        
                        <%End If%>
                        
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
                        <asp:Label ID="lbl_DO_CODE" runat="server" /></font>
                </td>
                <td>
                    <font size="2">
                        <asp:Label ID="DO_CODE" runat="server" />
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_STATUS" runat="server" /></font>
                </td>
                <td colspan="3">
                    <font size="2">
                        <asp:Label ID="DO_STATUS" runat="server" /></font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_STORER_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="STORER_CODE" runat="server" AutoPostBack="true" />
                    </font>
                </td>
           <%--     <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TRACK_NO" runat="server" /></font>
                </td>
                <td >
                    <asp:TextBox ID="DO_TRACK_NO" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
                    <asp:HiddenField ID="hide_DO_TRACK_NO" runat="server" />
                </td>                --%>                      
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_EDI_WIT_NO" Text="WH ISS Ticket" runat="server" /></font>
                </td>
                <td >
                    <font size="2">
                    <asp:TextBox ID="DO_CUS_REF_NO" MaxLength="20" runat="server" visible="false"/>
                    <asp:TextBox ID="DO_EDI_WIT_NO" runat="server" Width="150px" MaxLength="80" />
                    </font>
                </td>
            </tr>            
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_DATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="DO_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_DATE" PopupButtonID="ImageButton1" Format="dd/MM/yyyy" />   
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_ISSUED_BY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_ISSUED_BY" runat="server" MaxLength="20"></asp:TextBox>
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CO_CODE" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_CO_CODE" runat="server" MaxLength="20" ReadOnly="true"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                       <asp:Label ID="lbl_DO_CUS_REF_NO" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_EDI_SIR_NO" MaxLength="20" runat="server" />
                </td>   
            
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TARGET_DELDATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="DO_TARGET_DELDATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender2" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_TARGET_DELDATE" PopupButtonID="ImageButton2" Format="dd/MM/yyyy" />
                </td>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CONF_DELDATE" runat="server" /></font>
                </td>
                <td width="150px">
                    <asp:TextBox ID="DO_CONF_DELDATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender3" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_CONF_DELDATE" PopupButtonID="ImageButton3" Format="dd/MM/yyyy" />      
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CONF_DELTIME" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_CONF_DELTIME" runat="server" Width="50" MaxLength="5"></asp:TextBox>
                </td>                
            </tr>
            <tr>
                <td class="LabelTD" colspan="0" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_ARRIVAL_DATE" runat="server" /></font>
                </td>
                <td width="150px" colspan="5">
                    <asp:TextBox ID="DO_ARRIVAL_DATE" runat="server" Width="80" MaxLength="10" onkeypress="return maskDate(event);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="../../images/calendar1.gif" 
                            ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                    <asp:CalendarExtender ID="CalendarExtender4" runat="server" CssClass="ajax_calendar" 
                            TargetControlID="DO_ARRIVAL_DATE" PopupButtonID="ImageButton4" Format="dd/MM/yyyy" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_FTRACK_NO" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DO_FTRACK_NO" runat="server" MaxLength="20" Width="210px"></asp:TextBox>
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TRANS_TYPE" runat="server" /></font>
                </td>
                <td>
                    <asp:DropDownList ID="DO_TRANS_TYPE" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TROLLEY_ID" Text="Trolley ID" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_TROLLEY_ID" runat="server" MaxLength="80" Width="120px" />
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_DRUM_ID" Text="Drum ID" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DO_DRUM_ID" runat="server" MaxLength="80" Width="120px" />
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_PROJECT_NO" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="DO_PROJECT_NO" runat="server" MaxLength="20" 
                        AutoPostBack="True"></asp:DropDownList>
                    </font>
                </td>
                <td colspan="4"><font size="2">
                    <asp:Label ID="lbl_PRJ_NAME" runat="server" />
                </font></td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_SENDER" runat="server" MaxLength="100" ></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER_TEL" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DO_SENDER_TEL" runat="server" MaxLength="100"></asp:TextBox>
                </td>         
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER_ADDR" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_SENDER_ADDR" runat="server" TextMode="MultiLine" Rows="3" Width="600px"></asp:TextBox>
                </td>         
            </tr>

              <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER_REGION" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_SENDER_REGION" runat="server" MaxLength="100" ></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER_PROVINCE" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_SENDER_PROVINCE" runat="server" MaxLength="100" ></asp:TextBox>
                </td>          
                  
                    <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SENDER_COUNTRY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_SENDER_COUNTRY" runat="server" MaxLength="100" ></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_CUS_CODE" runat="server" /></font>
                </td>
                <td runat="server">
                    <font size="2">
                    <asp:DropDownList ID="CUS_CODE" runat="server" MaxLength="20" 
                        AutoPostBack="True"></asp:DropDownList>
                    </font>
                </td>
                <td colspan="4"><font size="2">
                    <asp:TextBox ID="CUS_NAME" runat="server" MaxLength="80" Width="200px" />
                </font></td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap rowspan="5">
                    <font size="2">
                        <asp:Label ID="lbl_DO_ADDR1" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_ADDR1" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_ADDR2" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_ADDR3" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            
            <tr>
                <!--
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_AREA_DEL" runat="server" /></font>
                </td>
                -->
                <td>
                    <asp:TextBox ID="DO_AREA_DEL" runat="server" MaxLength="50"></asp:TextBox>
                </td>
                <!--
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_REGION_DEL" runat="server" /></font>
                </td>
                -->
                <td colspan="4">
                    <asp:TextBox ID="DO_REGION_DEL" runat="server" MaxLength="50"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <!--
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_COUNTRY_DEL" runat="server" /></font>
                </td>
                -->
                <td colspan="5">
                    <asp:TextBox ID="DO_COUNTRY_DEL" runat="server" MaxLength="50"></asp:TextBox>
                </td>    
            </tr>
            <tr>
               <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_PROVINCE" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox runat="server" ID="DO_PROVINCE" MaxLength="200" />
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CITY" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox runat="server" ID="DO_CITY" MaxLength="200" />
                </td>   
                
                 <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_ROUTE_ID" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="ROUTE_ID" runat="server" MaxLength="100" ></asp:TextBox>
                </td>
            </tr>          
           

           <!-- new add --> 
           <tr><td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_DO_INV_NO" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_INV_NO" runat="server" Width="200" MaxLength="50"></asp:TextBox>
                </td>  
           </tr>
           <tr><td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_DO_CONSIGNEE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_CONSIGNEE" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
           <tr>
                <td class="LabelTD" nowrap rowspan="3">
                    <font size="2">
                        <asp:Label ID="lbl_DO_CONSIGNEE_ADDR1" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_CONSIGNEE_ADDR1" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>

            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_CONSIGNEE_ADDR2" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_CONSIGNEE_ADDR3" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_DO_POST_CODE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_POST_CODE" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>
            </tr>

            <tr>
                <td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_DO_DELIVERY_RMKS" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_DELIVERY_RMKS" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                        <font size="2">
                        <asp:Label ID="lbl_DO_SHIP_TO" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_SHIP_TO" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
           <tr>
                <td class="LabelTD" nowrap rowspan="3">
                    <font size="2">
                        <asp:Label ID="lbl_DO_SHIP_ADDR1" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_SHIP_ADDR1" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>

            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_SHIP_ADDR2" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td colspan="5">
                    <asp:TextBox ID="DO_SHIP_ADDR3" runat="server" Width="400" MaxLength="100"></asp:TextBox>
                </td>  
            </tr>
            <!--  end  --->
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CUS_CONT" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_CUS_CONT" runat="server" MaxLength="50"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_CUS_CONT_TEL" runat="server" /></font>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="DO_CUS_CONT_TEL" runat="server" MaxLength="20"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_DRIVER" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_DRIVER" runat="server" MaxLength="20"></asp:TextBox>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_DRIVER_TEL" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_DRIVER_TEL" runat="server" MaxLength="20"></asp:TextBox>
                </td>  
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_VEHICLE_NO" runat="server" /></font>
                </td>
                <td>
                    <asp:TextBox ID="DO_VEHICLE_NO" runat="server" MaxLength="20"></asp:TextBox>
                </td>  
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TOTL_PALLETS" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="DO_TOTL_PALLETS" runat="server" MaxLength="12" style="text-align:right"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TOTL_CARTONS" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="DO_TOTL_CARTONS" runat="server" MaxLength="12" style="text-align:right"></asp:TextBox>
                    </font>
                </td>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TOTL_BINS" runat="server" /></font>
                </td>
                <td nowrap>
                    <font size="2">
                        <asp:TextBox ID="DO_TOTL_BINS" runat="server" MaxLength="12" style="text-align:right"></asp:TextBox>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_REM" runat="server" Height="117px" Width="384px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TO_STORER_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_TO_STORER_REM" runat="server" Height="117px" Width="384px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_STORER_REM" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_STORER_REM" runat="server" Height="117px" Width="384px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <!------ NEW ADD ------->
           <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_PAY_TERMS" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_PAY_TERMS" runat="server" Width="384px" TextMode="MultiLine" Rows="3"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_TRADE_TERMS" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:TextBox ID="DO_TRADE_TERMS" runat="server" Width="384px" TextMode="MultiLine" Rows="3"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="LabelTD" nowrap>
                    <font size="2">
                        <asp:Label ID="lbl_DO_SHIP_MODE" runat="server" /></font>
                </td>
                <td colspan="5">
                    <asp:DropDownList ID="DO_SHIP_MODE" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <!----- END ----------->
            
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

            <br />
                 <%--   <table  width="100%">
                        <tr>
                            <td class="TITLE" width="50%">
                                Packing Video List
                            </td>
                            <td class="TITLE">
                                Tracking No. List:
                            </td>
                        </tr>
                        <tr>
                            <td width="50%">
                                <asp:GridView ID="GridView2" runat="server" Height="10px" Width="640px" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="dov_seq" HorizontalAlign="Left">
                                <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Video Seq" HeaderStyle-Width="70px">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="dov_seq" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Video Name">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label runat="server" id="DOV_VIDEO_NAME_SHOW" Visible="false" />
                                            <asp:LinkButton runat="server" ID="DOV_VIDEO_NAME" CommandName="DLVIDEO" />
                                            <br />
                                            <asp:Label runat="server" ID="lbl_sys_cd" Font-Size="Smaller" Text="Video Uploaded Date:" /><asp:label runat="server" ID="SYS_CD" Font-Size="Smaller" Font-Italic="true" /><br />
                                            <asp:Label runat="server" ID="lbl_sys_cb" Font-Size="Smaller" Text="Video Uploaded By:" /><asp:label runat="server" ID="SYS_CB" Font-Size="Smaller" Font-Italic="true" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Box Weight">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="DOV_WEIGHT" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Box No.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:Label ID="DOV_BOX_NO" runat="server" Font-Size="11px" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:button runat="server" ID="btnDel" CssClass="all_button" Text="Delete Video" CommandName="DELVIDEO" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <asp:button runat="server" ID="btnPrint" CssClass="all_button" Text="Print Label" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                                </asp:GridView>
                            </td>
                            <td valign="top">
                                <asp:GridView ID="GridView3" runat="server" Height="10px" Width="640px" Font-Names="Arial"
                                Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Tracking Record Found."
                                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                                CellPadding="3" CaptionAlign="Top" DataKeyNames="DOB_SEQ" HorizontalAlign="Left">
                                <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                                <Columns>
                                    <asp:BoundField runat="server" DataField="DOB_SEQ" HeaderText="Tracking Seq" ItemStyle-Font-Size="11px" />
                                    <asp:BoundField runat="server" DataField="DOB_BOX_NO" HeaderText="Box No." ItemStyle-Font-Size="11px"/>
                                    <asp:TemplateField HeaderText="Tracking NO.">
                                        <ItemTemplate>
                                            <asp:Linkbutton runat="server" id="DOB_FTRACK_NO" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField runat="server" DataField="DOB_SUB_FTRACK_NO" HeaderText="Sub Ftrack No." ItemStyle-Font-Size="11px"/>
                                    <asp:BoundField runat="server" DataField="DOB_STATUS" HeaderText="Status" ItemStyle-Font-Size="11px"/>                                    
                                    <asp:BoundField runat="server" DataField="DOB_RMKS" HeaderText="Remarks" ItemStyle-Font-Size="11px"/>
                                </Columns>
                                </asp:GridView>
                            </td>
                        </tr>
                    </table>
                <br />--%>

        <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 900px">
            <%--<tr>
                <td colspan="6" class="menuTD">
                &nbsp;
                    <asp:Button ID="newrow" runat="server" Text="Add" CssClass="all_button" Visible="false" />
                    <asp:Button ID="selectItemBtn" runat="server" Text="Select Item" CssClass="all_button" UseSubmitBehavior ="false" />
                </td>
            </tr>--%>
            <tr>
                <td colspan="6">   
                    <asp:UpdatePanel runat="server" ID="updtPnl_detailsList" UpdateMode="Conditional" RenderMode="Inline">
            <ContentTemplate>                    
                    <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                        Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" EmptyDataText="No Record Found."
                        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="3" CaptionAlign="Top" DataKeyNames="dod_seq" HorizontalAlign="Left" ShowFooter="true">
                        <Columns>      
                            <asp:TemplateField HeaderText="Seq No.">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_disp_seq" width="50" MaxLength="10" runat="server" Font-Size="11px" style="text-align:center" ></asp:TextBox>
                                    <asp:HiddenField ID="itm_serial_no_yn" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Subinventory">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_WH_CODE" runat="server" Font-Size="11px" Width="50px" MaxLength="20" AutoPostBack="true" OnTextChanged="DOD_WH_CODE_TextChanged"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>

                             <asp:TemplateField HeaderText="Floor">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="dod_fl_code" runat="server" Font-Size="11px" OnSelectedIndexChanged="DOD_FL_CODE_SelectedIndexChanged" AutoPostBack="true"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                              <asp:TemplateField HeaderText="Location">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:dropdownlist runat="server" ID="DOD_LOC_WH" Font-Size="11px" />
                                </ItemTemplate>                          
                            </asp:TemplateField>

                            <%-- <asp:TemplateField HeaderText="Location">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_LOC_WH" runat="server" Font-Size="11px" Width="50px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>--%>

                            <asp:TemplateField HeaderText="DO No.">
                                <ControlStyle Width="50px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_TICKET_NO" runat="server" Font-Size="11px" Width="50px" MaxLength="20"></asp:TextBox>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Pallet No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_pallet_no" width="70" MaxLength="20" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Carton No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_carton_no" width="70" MaxLength="20" runat="server" Font-Size="11px" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Pack No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="dod_pack_no" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Lot">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_batch_no" runat="server" Font-Size="11px" Width="70px"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>

                             <asp:TemplateField HeaderText="Expiry Date" Visible="false">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_manu_date" width="70" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>

                               <asp:TemplateField HeaderText="Last Lot No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_LAST_LOT" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                             <asp:TemplateField HeaderText="Max Lot No.">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_MAX_LOT" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Vendor Code">
                                <ControlStyle Width="80px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_vnd_code" width="70" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>  
                            
                            <asp:TemplateField HeaderText="Internal Item Code WMS">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_itm_code" width="150" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>  

                            <asp:TemplateField HeaderText="Item Code">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="itm_sku_no" width="150" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField> 
                            
                            <asp:TemplateField HeaderText="Item Name">
                                <ControlStyle Width="200px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_itm_desc" width="80" runat="server" Font-Size="11px"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <%--<asp:TemplateField HeaderText="Cut" Visible="false">
                                <ControlStyle Width="30px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:CheckBox ID="dod_cut_yn" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>--%>

                             <asp:TemplateField HeaderText="Return Qty">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>                                     
                                    <asp:TextBox ID="dod_return_qty" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Pack Type">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="dod_pack_type" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Pack Key">
                                <ControlStyle Width="40px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_pack_key" width="70" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Expiry Date">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:Label ID="dod_expiry_date" width="70" runat="server" Font-Size="11px" />
                                </ItemTemplate>
                            </asp:TemplateField>                            
                            
                            <asp:TemplateField HeaderText="SS QTY">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_SS_QTY" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="PK QTY">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="DOD_PK_QTY" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>                             
                           
                            <asp:TemplateField HeaderText="Qty">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_qty" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>

                               <asp:TemplateField HeaderText="Min Prod Date">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="DOD_MIN_PROD_DATE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                    <%-- <asp:ImageButton ID="btnCal01" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calEXP" runat="server" CssClass="ajax_calendar" TargetControlID="DOD_MIN_PROD_DATE" PopupButtonID="btnCal01" Format="dd-MMM-yyyy" />--%>
                                </ItemTemplate>
                            </asp:TemplateField>   
                            
                            <asp:TemplateField HeaderText="Min Shelf Life">
                                <ControlStyle />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Wrap="false" />
                                <ItemTemplate>
                                     <asp:TextBox ID="DOD_MIN_SHELF_LIFE" runat="server" Width ="80" MaxLength="10" Font-Size="11px" onkeypress="return maskNumOnly(event);"></asp:TextBox>
                                    <%-- <asp:ImageButton ID="btnCal02" runat="server" ImageUrl="../../images/calendar1.gif" ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                                     <asp:CalendarExtender ID="calEXP1" runat="server" CssClass="ajax_calendar" TargetControlID="DOD_MIN_SHELF_LIFE" PopupButtonID="btnCal02" Format="dd-MMM-yyyy" />--%>
                                </ItemTemplate>
                            </asp:TemplateField>                        

                            <asp:TemplateField HeaderText="UOM">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="dod_uom" runat="server" Font-Size="11px"  Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Number per UOM">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_pcs_uom" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Qty 2">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_qty2" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" onkeypress="return maskNumOnly(event);"/>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="UOM2">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="dod_uom2" runat="server" Font-Size="11px"  Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Total Number">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_totpcs" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Total Weight(kg)">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_tot_wgt" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Total Volume(cbm)">
                                <ControlStyle Width="70px" />
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_tot_cbm" width="70" MaxLength="12" runat="server" Font-Size="11px" style="text-align:right" />
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Remarks">
                                <ControlStyle Width="100px" />
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemTemplate>
                                    <asp:TextBox ID="dod_rem" width="70" MaxLength="200" runat="server" Font-Size="11px" />
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
                <td colspan="8" class="menuTD">
			<table width="100%">
			<tr><td class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" OnClientClick="return confirm(&quot;Save Recrod?&quot;);" />
                    <input id="btnBack" runat="server" type="button" value="Back" onclick="Javascript:window.location='../../cms_search.aspx?menu_code=OB_DO'"
                        class="all_button" />
                    <input type="button" runat="server" id="btnClose2" value="Close" onclick="Javascript: window.close();" class="all_button" />
			</td><td align="right" class="menuTD"><asp:Button ID="CancelBtn" Text ="Cancel" CssClass="all_button" runat ="server" /></td></tr>
			</table>
                </td>
            </tr>
        </table>
    </div>
            
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

        <asp:HiddenField ID="DO_CODE_HF" runat="server" />
                
        <asp:HiddenField ID="COList" runat ="server" />
        <asp:HiddenField ID="COSeqList" runat ="server" />
        <asp:HiddenField ID="itemList" runat ="server" />
        <asp:HiddenField ID="packKeyList" runat ="server" />        
        <asp:HiddenField ID="DO_PACK_LABEL_1" runat ="server" />
        <asp:HiddenField ID="DO_PACK_LABEL_2" runat ="server" />
        <asp:HiddenField ID="DO_PACK_LABEL_3" runat ="server" />
        <asp:HiddenField ID="DO_PACK_LABEL_QTY_1" runat ="server" />
        <asp:HiddenField ID="DO_PACK_LABEL_QTY_2" runat ="server" />
        <asp:HiddenField ID="DO_PACK_LABEL_QTY_3" runat ="server" />
        <asp:HiddenField ID="qtyList" runat ="server" />

        <asp:HiddenField ID="moduleAction" runat ="server" />

        <asp:HiddenField ID="INIT_JS" runat ="server" />
    </ContentTemplate>
    </asp:UpdatePanel>

    <asp:HiddenField ID="JS_MODULE_ID" Value="DO_MAIN" runat ="server" />

    <asp:UpdatePanel runat="server" ID="updtPnlLoad">
    <ContentTemplate>
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
    </ContentTemplate>
    </asp:UpdatePanel>

    <asp:HiddenField ID="loadDummy" runat="server" />
    <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
                TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
                DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
                Y="250">
            </asp:ModalPopupExtender>

    </form>
    <form name="hiddenForm" id="hiddenForm" method="post"/>
</body>
</html>
