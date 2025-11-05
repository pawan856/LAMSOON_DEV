
<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cms_search.aspx.vb" Inherits="cms_search" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<%@ Register Assembly="EO.Web" Namespace="EO.Web" TagPrefix="eo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <script language ="javascript" src="js/validation.js"></script>
    <script language ="javascript" src="js/JS_Calendar.js"></script>
    <script language ="javascript" src="js/mm.js"></script>
    <script language="javascript" src="js/listUtil.js"></script>
    <script language="javascript" src="js/formatUtil.js"></script>
    <script language="javascript" src="js/formPostInterfacing.js"></script>
    <script src="js/jquery-1.9.1.min.js"></script>
    <script src="js/jquery-migrate-1.1.1.min.js"></script> 
    <link rel="stylesheet" href="stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="stylesheet/general.css" type="text/css" />    
    
        <script type="text/javascript">
            function ChgChkState(id, checkState) {
                var cb = document.getElementById(id);
                if (cb != null)
                    cb.checked = checkState;
            }

            function ChgAllChkState(checkState) {
                if (chkArray != null) {
                    for (var i = 0; i < chkArray.length; i++)
                        ChgChkState(chkArray[i], checkState);
                }

                //var wholelist = document.getElementById('cbList');
                //var list = document.getElementById('cbSelected');

                //if (checkState) {
                //    list.value = wholelist.value;
                //} else {
                //    list.value = "";
                //}

                if (checkState) {
                    document.getElementById('toggle_selectAll').value = "Y";
                    document.getElementById('cbSelected').value = '';
                    document.getElementById('cbUnSelected').value = '';
                } else {
                    document.getElementById('toggle_selectAll').value = "N";
                    document.getElementById('cbUnSelected').value = '';
                    document.getElementById('cbSelected').value = '';
                }
            }

            function chkSelected(obj, val) {
                var list = document.getElementById('cbSelected');
                var ulist = document.getElementById('cbUnSelected');

                if (obj.checked) {
                    if (val != null) {
                        if (list.value == null || list.value == "") {
                            list.value = val;
                        } else if (list.value.indexOf(val) < 0) {
                            list.value += "," + val;
                        }
                        if (ulist.value.indexOf(val + ",") >= 0) {
                            ulist.value = ulist.value.replace(val + ",", "");
                        } else if (ulist.value.indexOf(val) >= 0) {
                            ulist.value = ulist.value.replace(val, "");
                        }
                    }
                } else {
                    if (val != null) {
                        if (ulist.value == null || ulist.value == "") {
                            ulist.value = val;
                        } else if (ulist.value.indexOf(val) < 0) {
                            ulist.value += "," + val;
                        }
                    }
                    if (list.value.indexOf(val + ",") >= 0) {
                        list.value = list.value.replace(val + ",", "");
                    } else if (list.value.indexOf(val) >= 0) {
                        list.value = list.value.replace(val, "");
                    }
                }
            }

            function Onscrollfnction() {
                var div = document.getElementById('mGVDiv');
                var div2 = document.getElementById('mGVHDDiv');
                //****** Scrolling HeaderDiv along with DataDiv ******
                div2.scrollLeft = div.scrollLeft;
                return false;
            }

            function CreateGridHeader(DataDiv, GridView1, HeaderDiv) {
                var DataDivObj = document.getElementById(DataDiv);
                var DataGridObj = document.getElementById(GridView1);
                var HeaderDivObj = document.getElementById(HeaderDiv);

                var HeadertableObj = HeaderDivObj.appendChild(document.createElement('table'));
                DataDivObj.style.paddingTop = '0px';
                var DataDivWidth = DataDivObj.clientWidth;

                HeaderDivObj.className = DataDivObj.className;
                HeaderDivObj.style.cssText = DataDivObj.style.cssText;

                HeaderDivObj.style.overflow = 'auto';

                HeaderDivObj.style.overflowX = 'hidden';

                HeaderDivObj.style.overflowY = 'hidden';
                HeaderDivObj.style.height = DataGridObj.rows[0].clientHeight + 'px';
                HeaderDivObj.style.borderBottomWidth = '0px';
                HeadertableObj.className = DataGridObj.className;
                HeadertableObj.style.cssText = DataGridObj.style.cssText;
                HeadertableObj.border = '1px';
                HeadertableObj.rules = 'all';
                HeadertableObj.cellPadding = DataGridObj.cellPadding;
                HeadertableObj.cellSpacing = DataGridObj.cellSpacing;

                var Row = HeadertableObj.insertRow(0);
                Row.className = DataGridObj.rows[0].className;
                Row.style.cssText = DataGridObj.rows[0].style.cssText;
                Row.style.fontWeight = 'bold';

                for (var iCntr = 0; iCntr < DataGridObj.rows[0].cells.length; iCntr++) {

                    if (DataGridObj.rows.length > 0) {
                        var spanTag = Row.appendChild(document.createElement('td'));
                        spanTag.innerHTML = DataGridObj.rows[0].cells[iCntr].innerHTML;

                        var width = 0;

                        if (DataGridObj.rows.length > 1) {
                            spanTag.style.fontWeight = 'bold';
                            spanTag.className = 'DtlLabel';
                            width = DataGridObj.rows[1].cells[iCntr].clientWidth;
                            DataGridObj.rows[1].cells[iCntr].style.width = width + 15 + 'px';
                        } else {
                            Row.style.fontWeight = 'normal';
                        }


                        if (iCntr <= DataGridObj.rows[0].cells.length - 2) {
                            spanTag.style.width = width + 15 + 'px';
                        }
                        else {
                            spanTag.style.width = width + 15 + 'px';
                        }
                    }
                }
                var tableWidth = DataGridObj.clientWidth;

                if (DataGridObj.rows[0] != null) {
                    DataGridObj.rows[0].style.display = 'none';
                }


                HeaderDivObj.style.width = DataDivWidth + 15 + 'px';
                DataDivObj.style.width = DataDivWidth + 15 + 'px';
                DataGridObj.style.width = tableWidth + 15 + 'px';

                HeadertableObj.style.width = tableWidth + 16 + 'px';
                return false;
            }



            function confirmPWD() {
                //document.myform.moduleAction.value = "CONFIRM_UNDO";
                document.srchForm.submit();
                alert('Action Undo Successfully');
            }
        </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="srchForm" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" AsyncPostBackTimeout="3600" />
    <eo:ScriptManager ID="eo_sm" runat="server">        
    </eo:ScriptManager>
    <asp:UpdatePanel ID="search_updt_panel" runat="server">
        <ContentTemplate>
            <table cellpadding="0" cellspacing="0" border="0" style="width: 100%">
                <tr>
                    <td style="background-color: transparent">
                        <div id="hddiv" runat="server">
                        </div>
                    </td>
                </tr>
                <tr>
                    <td style="background-color: transparent">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        <div id="mdiv" runat="server" visible="false" style="width: 100%; position: relative">
                        </div>
                        <div id="dtdiv" runat="server">
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div id="footerDiv" runat="server" visible="false">
                            <table id="Table1" cellpadding="0" cellspacing="0" width="100%" border="0" runat="server">
                                <tr>
                                    <td class="TITLE">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
            </table>
                    <asp:HiddenField ID="RowsCount" runat="server" />    
            <asp:Literal ID="CBCollection" runat="server"></asp:Literal>
            <asp:PlaceHolder runat="server" ID="hdfieldHolder"></asp:PlaceHolder>
            <asp:PlaceHolder runat="server" ID="customCtrlHolder"></asp:PlaceHolder>
            <asp:HiddenField ID="loadDummy" runat="server" />
            <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none"
                ScrollBars="None">
                <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
                    height: 80px">
                    <tr>
                        <td align="center" style="background-color: White; width: 100%; height: 40px; vertical-align: bottom">                            
                            <font color="#193B65" style="width: 100%; text-align: center; font-size: 16px;">Loading...please wait</font>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" style="background-color: White; width: 100%; vertical-align: top"
                            id="tdLoad">                            
                            <img src="images/wait_loading.gif" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>    
            <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
                TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground_transparent"
                DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
                Y="250">
            </asp:ModalPopupExtender>            
        </ContentTemplate>
    </asp:UpdatePanel> 
  
    </form>            
</body>
</html>

<script type="text/javascript" language="javascript">
    $(window).unload(function () {
        PageMethods.removeTemplateSession();
    });    
 </script>