<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cable_label_print.aspx.vb" Inherits="INBOUND_GR_GR_LABELS_item_label_print" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Print Item Label</title>
<link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
<link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />
<script language ="javascript" src="../../js/validation.js"></script>
<script language ="javascript" src="../../js/JS_Calendar.js"></script>
<script language="javascript" src="../../js/listUtil.js"></script>
<script language="javascript" src="../../js/formatUtil.js"></script>
<script language="javascript" src="../../js/formPostInterfacing.js"></script>
<script language="javascript">

    function PrintLabel(STORER_CODE, itemStr) {
            removeAllElementFromForm(document.hiddenForm);

            window.open("", "printLabel", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=800,height=600,left=5,top=15");

            setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", STORER_CODE);
            setInterfaceDataToForm(document.hiddenForm, "itemStr", itemStr);
            setInterfaceDataToForm(document.hiddenForm, "lblType", "ITEM_S");
            //setInterfaceDataToForm(document.hiddenForm, "LabelSize", document.getElementById("LabelSize").value);    
            document.hiddenForm.action = "../../REPORT/ITM_LABEL/item_label.aspx";
            document.hiddenForm.target = "printLabel";
            document.hiddenForm.submit();
     
    }

    function PrintLabelCable(STORER_CODE, itemStr) {
        removeAllElementFromForm(document.hiddenForm);

        window.open("", "printLabelC", "titlebar=0,location=0,menubar=0,toolbar=0,scrollbars=yes,personalbar=0,resizable=yes,status=1,width=800,height=600,left=30,top=75");

        setInterfaceDataToForm(document.hiddenForm, "STORER_CODE", STORER_CODE);
        setInterfaceDataToForm(document.hiddenForm, "itemStr", itemStr);
        setInterfaceDataToForm(document.hiddenForm, "lblType", "ITEM_S");
        //setInterfaceDataToForm(document.hiddenForm, "LabelSize", document.getElementById("LabelSize").value);    
        document.hiddenForm.action = "../../REPORT/ITM_LABEL/item_labelCable.aspx";
        document.hiddenForm.target = "printLabelC";
        document.hiddenForm.submit();

    }
</script>
</head>
<body>
    <form id="myform" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
     <table width="100%" bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
        <tr>
            <td class="TITLE">
            <asp:UpdatePanel runat="server" ID="updtPnlAlert">
                <ContentTemplate>
                    <asp:Button runat="server" ID="btnPrint" Text="Print" class="all_button" />
                    <input type="button" value="Close" id="btnClose" onclick="javascript:window.close();" class="all_button" />
                </ContentTemplate>
            </asp:UpdatePanel>
                
              
            </td>
        </tr>
        <tr>
            <td class="menuTD">
                <asp:Button runat="server" Text="Select All" ID="btnSelAll" CssClass="all_button" />
                <asp:Button runat="server" Text="Un-Select All" ID="btnUnAll" CssClass="all_button" />
                <div style="float:right">
                    <asp:DropDownList runat="server" ID="LabelSize" Visible="false">
                        <asp:ListItem Text="Normal" Value="N" />
                        <asp:ListItem Text="Small" Value="S" />
                        <asp:ListItem Text="A4" Value="A4" />
                    </asp:DropDownList>
                </div>
            </td>
        </tr>
        <tr>
            <td>
                <asp:GridView ID="GridView1" runat="server" Height="10px" Width="100%" Font-Names="Arial"
                     Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False" ShowHeaderWhenEmpty="true" 
                     BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" EmptyDataText="No Item record found!"
                     CellPadding="3" CaptionAlign="Top" DataKeyNames="ILBS_SEQ" HorizontalAlign="Left">
                      <RowStyle CssClass="GV" />
                      <HeaderStyle CssClass="DtlLabel" Font-Bold="False" />
                      <Columns>
                         <asp:TemplateField >
                            <ControlStyle Width="30px" />
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                            <ItemTemplate>
                                <asp:CheckBox runat="server" ID="nSelect" />
                                <asp:HiddenField runat="server" ID="ILOC_SEQ" />
                                <asp:HiddenField runat="server" ID="ILBS_SEQ" />                                
                                <asp:HiddenField runat="server" ID="ITM_TYPE" />
                            </ItemTemplate>
                         </asp:TemplateField>
                         <asp:BoundField runat="server" DataField="ITM_SKU_NO" HeaderText="Stock No." HeaderStyle-HorizontalAlign="Left" />
                         <asp:BoundField runat="server" DataField="ITM_NAME" HeaderText="Item Name" HeaderStyle-HorizontalAlign="Left" />
                         <asp:BoundField runat="server" DataField="ILBS_SERIAL_NO" HeaderText="Serial No" HeaderStyle-HorizontalAlign="Left" />
                         <asp:BoundField runat="server" DataField="ILBS_DRUM_ID" HeaderText="Drum ID" HeaderStyle-HorizontalAlign="Left" />
                         <asp:BoundField runat="server" DataField="ILBS_QTY2" HeaderText="Qty 2" HeaderStyle-HorizontalAlign="Right" />
                         <asp:BoundField runat="server" DataField="ILBS_UOM2" HeaderText="UOM" HeaderStyle-HorizontalAlign="Left" />
                         
                         
                      </Columns> 
                </asp:GridView>
            </td>
        </tr> 
        <tr>
            <td class="TITLE">
                <asp:UpdatePanel runat="server" ID="UpdatePanel1">
                <ContentTemplate>
                <asp:Button runat="server" ID="btnPrint2" Text="Print" class="all_button" />
                <input type="button" value="Close" id="btnClose2" onclick="javascript:window.close();" class="all_button"/>
                </ContentTemplate>
                </asp:UpdatePanel> 
                
            </td>
        </tr>       
      </table>
    </div>
    </form>
<form name="hiddenForm" id="hiddenForm" method="post" />    
</body>
</html>
