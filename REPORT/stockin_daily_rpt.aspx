<%@ Page Language="VB" AutoEventWireup="false" CodeFile="stockin_daily_rpt.aspx.vb" Inherits="REPORT_stockin_daily_rpt" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style>
        TD.GV
            {
                font-family: "Arial" , "Helvetica" , "sans-serif";
                font-size: 11px;
                /*background-color: #d4d0c8;*/
                background-color: #F0F0F0;
            }
    
        .TITLE
            {
                font-family: "Arial" , "Helvetica" , "sans-serif";
                font-size: 16px;
                /*background-color: #006699;*/
                color: #FFFFFF;
                /*color: #4C60B6;*/
                background-color: #92A1B9;
            }   
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
        
        <table cellpadding="2" cellspacing="0" width="100%" border="0" style="background-color: #EEEEEE">
            <tr>
                <td width="80%">
                    <rsweb:ReportViewer ID="ReportViewer1" runat="server" font-size="8pt" height="600px" width="100%" CssClass=""></rsweb:ReportViewer>
                </td>
            
            <td width="20%" style="vertical-align: top; border: 1px; border-style:solid">
                <table cellpadding="0" cellspacing="0" style="border-style: solid; border: 1px;" width="100%">
                    <tr><td class="TITLE" colspan="2">GR Code Filter</td></tr>
                    <tr>
                        <td class="GV">
                        <asp:DataList ID="DataList1" runat="server">
                                <ItemTemplate>
                                     <table>
                                        <tr style="vertical-align: bottom">
                                            <td style="vertical-align: bottom"><asp:CheckBox runat="server" ID="gr_code_cb" Checked="true" Text='<%#DataBinder.Eval(Container.DataItem, "gr_code")%>' />&nbsp;&nbsp;-&nbsp;&nbsp;
                                             </td>
                                            <td><asp:Label runat="server" ID="label1" Text='<%# DataBinder.Eval(Container.DataItem,"gr_ref_no") %>'></asp:Label></td>
                                        </tr>
                                     </table>
                                     
                                </ItemTemplate>
                            </asp:DataList>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" class="GV">
                            <asp:Button runat="server" ID="btnFilter" Text="Filter Report" />
                        </td>
                    </tr>
                </table>
            </td>
            </tr>
        </table>
        
    </div>
    </form>
</body>
</html>
