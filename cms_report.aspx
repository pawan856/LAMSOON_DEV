<%@ Page Language="VB" AutoEventWireup="false" CodeFile="cms_report.aspx.vb" Inherits="cms_report" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language ="javascript" src="js/validation.js"></script>
<script language ="javascript" src="js/JS_Calendar.js"></script>
<script language ="javascript" src="js/mm.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
     <title>Report Generator</title>
    <link rel="stylesheet" href="stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="stylesheet/general.css" type="text/css" />    
    
     <script type="text/javascript">
        function ChgChkState(id, checkState)
        {
            var cb = document.getElementById(id);
            if (cb != null)
               cb.checked = checkState;
        }

        function ChgAllChkState(checkState)
        {
            if (chkArray != null)
            {
                for (var i = 0; i < chkArray.length; i++)
                    ChgChkState(chkArray[i], checkState);
            }
        }

        function ChgAllChkStateDT(checkState) {
            if (chkArray != null) {
                for (var i = 0; i < chkArrayDT.length; i++)
                    ChgChkState(chkArrayDT[i], checkState);
            }
        }
        
        function ChgHDState()
        {
            if (chkArray != null)
            {
                for (var i = 1; i < chkArray.length; i++)
                {
                    var cb = document.getElementById(chkArray[i]);
                    if (!cb.checked)
                    {
                        ChgChkState(chkArray[0], false);
                        return;
                    }
                }

                ChgChkState(chkArray[0], true);
            }
        }

        function ChgDTState() {
            if (chkArrayDT != null) {
                for (var i = 1; i < chkArrayDT.length; i++) {
                    var cb = document.getElementById(chkArrayDT[i]);
                    if (!cb.checked) {
                        ChgChkState(chkArrayDT[0], false);
                        return;
                    }
                }

                ChgChkState(chkArrayDT[0], true);
            }
        }
        
        function maskKey(objEvent) {
            var iKeyCode;
            iKeyCode = objEvent.keyCode;
            if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
            return false;
        }

    </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form method="post" name="rptform" id="rptform" runat="server" defaultbutton="Submit">
    <div id="hddiv" runat="server"></div>
        <div id="col_sel" runat="server">
        <br />
        <table border="0" cellspacing="1" cellpadding="1" align="center" width="820px">
            <tr>
                <td colspan="2" class="TITLE">
                    <b><asp:Label ID="lblHD" runat="server" /></b>
                </td>
            </tr>
        </table>
        <asp:GridView ID="gvrsList" runat="server" Height="10px" Width="820px"
            Font-Names="Arial" Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False"
            EmptyDataText="No Record Found." CaptionAlign="Top" BorderStyle="Solid"
            DataKeyNames="column_name" HorizontalAlign="Center">
            <PagerStyle Font-Size="10px"
                BorderColor="Transparent" BorderWidth="1px" Font-Names="Arial" />
            <Columns>
                <asp:TemplateField>
                    <HeaderStyle  Width="5%"/>
                    <HeaderTemplate>
                        <asp:CheckBox runat="server" ID="chkSelectAll" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="chkSelect" runat="server" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Sum">
                    <HeaderStyle HorizontalAlign="Center" Width="8%"/>
                    <ItemTemplate>
                        <asp:CheckBox ID="chkSum" runat="server" Enabled="false" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="Field Name">
                    <HeaderStyle HorizontalAlign="Left"  Width="37%" />
                    <ItemTemplate>
                        <asp:Label ID="cold_label" runat="server" Width="150px" Font-Size="10px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Seq">
                    <HeaderStyle HorizontalAlign="Left" Width="12%" />
                    <ItemTemplate>
                        <asp:TextBox ID="cold_display_seq" runat="server" Width="30px" Font-Size="10px" style="text-align:center" ></asp:TextBox>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Alignment">
                    <HeaderStyle HorizontalAlign="Left" Width="18%" />
                    <ItemTemplate>
                        <asp:DropDownList ID ="cold_align" runat="server"></asp:DropDownList>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Width">
                    <HeaderStyle HorizontalAlign="Left" Width="20%" />
                    <ItemTemplate>
                        <asp:TextBox ID="cold_width" runat="server"  Font-Size="10px"  Width="40px"></asp:TextBox><font size="1"><b> cm</b></font>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <RowStyle BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px"
                Font-Size="9px" />
            <HeaderStyle BorderColor="White" BorderStyle="Inset" BorderWidth="1px"
                Font-Names="Arial" Font-Size="11px" HorizontalAlign="Center"
                VerticalAlign="Middle" Wrap="True" />
        </asp:GridView>
        <table id="detailBar" border="0" cellspacing="1" cellpadding="1" align="center" width="820px" runat="server" visible="false">
            <tr>
                <td colspan="2" class="TITLE">
                   <b><asp:Label ID="lblDT" runat="server" /></b>
                </td>
            </tr>
        </table>
        <asp:GridView ID="gvrsList2" runat="server" Height="10px" Width="820px"
            Font-Names="Arial" Font-Overline="False" Font-Size="10px" AutoGenerateColumns="False"
            EmptyDataText="No Record Found." CaptionAlign="Top" BorderStyle="Solid"
            DataKeyNames="column_name" HorizontalAlign="Center">
            <PagerStyle Font-Size="10px"
                BorderColor="Transparent" BorderWidth="1px" Font-Names="Arial" />
            <Columns>
                <asp:TemplateField>
                    <HeaderStyle  Width="5%"/>
                    <HeaderTemplate>
                        <asp:CheckBox runat="server" ID="chkSelectAll" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="chkSelect" runat="server" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Sum">
                    <HeaderStyle HorizontalAlign="Center" Width="8%"/>
                    <ItemTemplate>
                        <asp:CheckBox ID="chkSum" runat="server" Enabled="false" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="Field Name">
                    <HeaderStyle HorizontalAlign="Left"  Width="37%" />
                    <ItemTemplate>
                        <asp:Label ID="cold_label" runat="server" Width="150px" Font-Size="10px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Seq">
                    <HeaderStyle HorizontalAlign="Left" Width="12%" />
                    <ItemTemplate>
                        <asp:TextBox ID="cold_display_seq" runat="server" Width="30px" Font-Size="10px" style="text-align:center" ></asp:TextBox>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Alignment">
                    <HeaderStyle HorizontalAlign="Left" Width="18%" />
                    <ItemTemplate>
                        <asp:DropDownList ID ="cold_align" runat="server"></asp:DropDownList>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Left" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Width">
                    <HeaderStyle HorizontalAlign="Left" Width="20%" />
                    <ItemTemplate>
                        <asp:TextBox ID="cold_width" runat="server"  Font-Size="10px"  Width="40px"></asp:TextBox><font size="1"><b> cm</b></font>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <RowStyle BorderColor="Silver" BorderStyle="Solid" BorderWidth="1px"
                Font-Size="9px" />
            <HeaderStyle BorderColor="White" BorderStyle="Inset" BorderWidth="1px"
                Font-Names="Arial" Font-Size="11px" HorizontalAlign="Center"
                VerticalAlign="Middle" Wrap="True" />
        </asp:GridView>
    </div>
    <div id="sort" runat="server">
     <table border="0" cellspacing="1" cellpadding="1" align="center" width="820px" id="sort_tbl" runat="server">
            <tr runat="server" id="sort_blank">
                <td colspan="2">
                   &nbsp;
                </td>
            </tr>
            <tr>
                <td id="sort_td" runat="server" class="LabelTD" style="width:20%">
                        <asp:Label ID="lblSort" runat="server" Text="Sorted By:" Font-Size="10pt"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="DDL_ORDERBY1" runat="server">
                    </asp:DropDownList>                  
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                  <asp:DropDownList ID="DDL_ORDERBY2" runat="server">
                    </asp:DropDownList>                    
                </td>
            </tr>         
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                  <asp:DropDownList ID="DDL_ORDERBY3" runat="server">
                    </asp:DropDownList>                    
                </td>
            </tr>                          
        </table>
    
    </div>
    
     
          
     <div id="file_dl" runat="server">
        <br />

            <table border="0" cellspacing="1" cellpadding="1" width="820px" align="center">
                <tr style="height: auto">
                    <td class="TITLE" colspan="3"><b><asp:Label ID="lbl_File"  Text="Report Detail" runat="server" /></b></td>
                </tr>
	            <tr style="height: 22px">
		            <td colspan="3" class="menuTD" runat="server" id="status_td">
                    <asp:Label ID="lbl_file_status" runat="server"  Style="position: static" 
                            Text='Please click the button below to Open / Download the report as your request.' 
                            ForeColor="#0000CC" Font-Size="10pt" width="500px"></asp:Label>
                    </td>
	            </tr>
	            <tr style="height: 22px" id="preview_tr" runat="server" visible="false">
	                <td class="LabelTD" valign="middle" style="width: 150px">
	                    <asp:Label ID="lbl_printout" runat="server" Text="Open Preview"></asp:Label>
	                    :
	                    </td>
	                <td colspan="2" style="width: 500px"><asp:ImageButton ID="btnPreview" runat="server" ImageUrl="images/preview.gif" 
	                onMouseOut="MM_swapImgRestore()" onMousedown="MM_swapImage('btnPreview','','images/preview_press.gif',1)" />
	                </td>
	            </tr>
	            <tr style="height: 22px">
	                <td class="LabelTD" valign="middle" style="width: 150px">
	                    <asp:Label ID="lbl_download" runat="server" Text="Print / Download As"></asp:Label>
	                    :
	                </td>
	                <td id="save_as_n_td" runat="server" style="width: 70px" visible="false"></td>
	                <td style="width: 500px"><asp:ImageButton ID="btnDownload" runat="server" ImageUrl="images/report_download.gif" 
	                onMouseOut="MM_swapImgRestore()" onMousedown="MM_swapImage('btnDownload','','images/report_download.gif',1)" />
	                </td>
	            </tr>
	            <tr style="height: auto">
		            <td class="TITLE" colspan="3">&nbsp;</td>
	            </tr>
            </table>
    </div> 
       
    <asp:Literal ID="CBCollection" runat="server"></asp:Literal>
    <asp:Literal ID="CBCollectionDT" runat="server"></asp:Literal>
    </form>
</body>
</html>
