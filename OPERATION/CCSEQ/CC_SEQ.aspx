<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CC_SEQ.aspx.vb" Inherits="OPERATION_CCSEQ_CC_SEQ" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Cycle Count Sequence Setup</title>
    <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

    <script language ="javascript" src="../../js/validation.js"></script>
    <script language ="javascript" src="../../js/JS_Calendar.js"></script>
    <script language="javascript" src="../../js/listUtil.js"></script>
    <script language="javascript" src="../../js/formatUtil.js"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js"></script>
    <script language="javascript">
        function getLoad() {
            var load_modalPopup = $find('load_ModalPopupExtender');
            load_modalPopup.show();
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" />
    <div>
    <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 1100px">
          <tr>
              <td width="100%" class="TITLE" colspan="3">
                  <b>
                      <asp:label ID="lheader" runat ="server" Text="Cycle Count Seq. Setup" /></b>
              </td>                    
              <td class="TITLE" align="right" >
                <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Back"
                          <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.location='../../user_group_fn.aspx?User_ID=&menu_code=MENU_OP'"
                          class="all_button" />
              </td>
          </tr>
          <tr>
             <td class="TITLE" colspan="4">
                 <font size="2"><asp:Label runat="server" ID="lbl_DLHD" Text="Download Template:" /></font>
             </td>
          </tr>
          <tr>
             <td class="LabelTD" width="150px">
                 <asp:Label runat="server" ID="lbl_DL_YEAR" Text="Year:" />
             </td>
             <td >
                 <asp:DropDownList runat="server" ID="DL_YEAR">
                    <asp:ListItem Value="2015" />
                    <asp:ListItem Value="2016" />
                    <asp:ListItem Value="2017" />
                    <asp:ListItem Value="2018" />
                    <asp:ListItem Value="2019" />
                    <asp:ListItem Value="2020" />
                    <asp:ListItem Value="2021" />
                    <asp:ListItem Value="2022" />
                    <asp:ListItem Value="2023" />
                    <asp:ListItem Value="2024" />
                    <asp:ListItem Value="2025" />
                    <asp:ListItem Value="2026" />
                    <asp:ListItem Value="2027" />
                    <asp:ListItem Value="2028" />
                    <asp:ListItem Value="2029" />
                    <asp:ListItem Value="2030" />
                    <asp:ListItem Value="2031" />
                    <asp:ListItem Value="2032" />
                    <asp:ListItem Value="2033" />
                    <asp:ListItem Value="2034" />
                    <asp:ListItem Value="2035" />
                 </asp:DropDownList>
             </td>
             <td class="LabelTD" width="150px">
                 <asp:Label runat="server" ID="lbl_DL_Period" Text="Period:" />
             </td>
             <td >
                 <asp:DropDownList runat="server" ID="DL_PERIOD">
                    <asp:ListItem Value="1YR" Text="1 Year" />
                    <asp:ListItem Value="2YR" Text="2 Year" />
                    <asp:ListItem Value="4YR" Text="4 Year" />
                    <asp:ListItem Value="ALL" Text="ALL" />
                 </asp:DropDownList>
             </td>
          </tr>
          <tr>
            <td class="LabelTD" width="150px">
                 <asp:Label runat="server" ID="lbl_DL_CCLS_WH" Text="Main Warehouse:" />
             </td>
             <td colspan="3">
                <asp:DropDownList runat="server" ID="DL_CCLS_WH" />
             </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td colspan="3">
                <asp:Button runat="server" ID="btnDL" text="Download Template Data" CssClass="all_button" OnClientClick="getLoad();" />
            </td>
          </tr>
          <tr>
             <td colspan="4" align="center" style=" background:#FFFFFF;">
               &nbsp;
            </td>
          </tr>
          <tr>
             <td class="TITLE" colspan="4">
                 <font size="2"><asp:Label runat="server" ID="lbl_UPHD" Text="Upload Data:" /></font>
             </td>
          </tr>
          <tr>
             <td class="LabelTD">
                 <asp:Label runat="server" ID="lbl_UP_YEAR" Text="Year:" />
             </td>
             <td>
                 <asp:DropDownList runat="server" ID="UP_YEAR">
                    <asp:ListItem Value="2015" />
                    <asp:ListItem Value="2016" />
                    <asp:ListItem Value="2017" />
                    <asp:ListItem Value="2018" />
                    <asp:ListItem Value="2019" />
                    <asp:ListItem Value="2020" />
                    <asp:ListItem Value="2021" />
                    <asp:ListItem Value="2022" />
                    <asp:ListItem Value="2023" />
                    <asp:ListItem Value="2024" />
                    <asp:ListItem Value="2025" />
                    <asp:ListItem Value="2026" />
                    <asp:ListItem Value="2027" />
                    <asp:ListItem Value="2028" />
                    <asp:ListItem Value="2029" />
                    <asp:ListItem Value="2030" />
                    <asp:ListItem Value="2031" />
                    <asp:ListItem Value="2032" />
                    <asp:ListItem Value="2033" />
                    <asp:ListItem Value="2034" />
                    <asp:ListItem Value="2035" />
                 </asp:DropDownList>
             </td>
             <td class="LabelTD">
                 <asp:Label runat="server" ID="lbl_UP_PERIOD" Text="Period:" />
             </td>
             <td >
                 <asp:DropDownList runat="server" ID="UP_PERIOD">
                    <asp:ListItem Value="1YR" Text="1 Year" />
                    <asp:ListItem Value="2YR" Text="2 Year" />
                    <asp:ListItem Value="4YR" Text="4 Year" />
                    <asp:ListItem Value="ALL" Text="ALL" />
                 </asp:DropDownList>
             </td>
          </tr>
          <tr>
            <td class="LabelTD" width="150px">
                 <asp:Label runat="server" ID="lbl_UP_CCLS_WH" Text="Main Warehouse:" />
             </td>
             <td colspan="3">
                <asp:DropDownList runat="server" ID="UP_CCLS_WH" />
             </td>
          </tr>
          <tr>
            <td>
                <asp:FileUpload runat="server" ID="UP_FILE" />
            </td>
            <td colspan="3">                
                <asp:Button runat="server" ID="btnUP" text="Upload Data" CssClass="all_button" OnClientClick="getLoad();" />
            </td>
          </tr>
          <tr>
            <td class="LabelTD">
                <asp:Label runat="server" ID="lbl_RST" Text="Upload Result:" />
            </td>
            <td colspan="3">
                <asp:TextBox runat="server" ID="txtUP_RESULT" TextMode="MultiLine" Rows="10" Width="800" />
            </td>
          </tr>
    </table>
    </div>

    <asp:Panel runat="server" CssClass="modalPopup" ID="loadPanel" Style="display: none" ScrollBars="None">
        <table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 250px;
            height: 80px">
            <tr>
                <td align="center" style="width: 100%; height: 80px; vertical-align: middle">                            
                    <font color="#193B65" style="width: 100%; text-align: center; font-size: 16px;">Processing data...please wait</font>
                </td>
            </tr>                    
        </table>
    </asp:Panel>
    <asp:HiddenField ID="loadDummy" runat="server" />
    <asp:ModalPopupExtender ID="load_ModalPopupExtender" runat="server" Enabled="True"
        TargetControlID="loadDummy" PopupControlID="loadPanel" BackgroundCssClass="modalBackground"
        DropShadow="false" RepositionMode="None" BehaviorID="load_ModalPopupExtender"
        Y="250">
    </asp:ModalPopupExtender>

    </form>
    <script type="text/javascript" language="javascript">

        
    </script>
</body>
</html>
