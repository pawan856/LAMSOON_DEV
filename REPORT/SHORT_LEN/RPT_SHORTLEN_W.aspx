<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RPT_SHORTLEN_W.aspx.vb" Inherits="REPORT_SHORT_LEN_RPT_SHORTLEN_W" %>
<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
     <title>RPT SHORT LENGTH</title>
  <link rel="stylesheet" href="../../stylesheet/newtext.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/button.css" type="text/css" />
    <link rel="stylesheet" href="../../stylesheet/general.css" type="text/css" />

    <script language ="javascript" src="../../js/validation.js"></script>
    <script language ="javascript" src="../../js/JS_Calendar.js"></script>
    <script language="javascript" src="../../js/listUtil.js"></script>
    <script language="javascript" src="../../js/formatUtil.js"></script>
    <script language="javascript" src="../../js/formPostInterfacing.js"></script>
    <script language="javascript">


    </script>

</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form id="form1" runat="server">
    <asp:ToolkitScriptManager ID="GlobalScriptMan" runat="server" EnablePageMethods="true" ScriptMode="Release" AsyncPostBackTimeout="18000" />
    <div>
       <table border="0" cellspacing="1" cellpadding="1" align="center" style="width: 800px">
        <tr>
           <td colspan="4" class="TITLE">
               <asp:label ID="lheader" text="Short Length Cable Summary Report" runat ="server" />                              
           </td>
        </tr>
        <tr>
            <td class="LabelTD">
             <asp:Label runat="server" Text="Start Date" ID="lbl_start_date" width="30%"/>
            </td>
            <td colspan="3">
                <asp:TextBox ID="START_DATE" runat="server" Width ="80" MaxLength="10" onkeypress="return maskDate(event);" CssClass="REQUIRED" />
                <asp:ImageButton ID="btnDATE_ID1" runat="server" ImageUrl="../../images/calendar1.gif" 
                        ImageAlign="Middle" BorderWidth="0" style="padding-left: 4px" />
                <asp:CalendarExtender ID="CalendarExtender1" runat="server" CssClass="ajax_calendar" 
                        TargetControlID="START_DATE" PopupButtonID="btnDATE_ID1" Format="dd/MM/yyyy" />
            </td>
        </tr>
        <tr>
            <td colspan="4" align="center">
                <asp:UpdatePanel runat="server" ID="UDP1" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:Button runat="server" ID="btnGen" CssClass="all_button" Text="Generate Report" />                        
                        <br />
                        <br />
                        <asp:Button runat="server" ID="btnDL" CssClass="all_button" Text="Download Report" Visible="false" />
                        <asp:hiddenfield runat="server" ID="hdf_file_path" />
                    </ContentTemplate>
                    <Triggers>
                        <asp:PostBackTrigger ControlID="btnDL" />
                    </Triggers>
                </asp:UpdatePanel>                                    
                <asp:UpdateProgress runat="server" ID="prog1"  DisplayAfter="0">
                    <ProgressTemplate>
                       <div style="width:100%; text-align:center;"> <asp:Image runat="server" ID="loadbar" ImageUrl="~/REPORT/SHORT_LEN/loadbar.gif" /> </div>
                    </ProgressTemplate>                
                </asp:UpdateProgress>                
            </td>
        </tr>
        <tr>
            <td colspan="4" align="center">
                <asp:UpdatePanel runat="server" ID="UDP2" RenderMode="Inline">
                    <ContentTemplate>
                        <asp:label runat="server" ID="errMsg" ForeColor="Red" />
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
        </table>
    </div>  
    </form>
</body>
</html>
