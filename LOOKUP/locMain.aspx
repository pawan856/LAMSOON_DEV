<%@ Page Language="VB" AutoEventWireup="false" CodeFile="locMain.aspx.vb" Inherits="LOOKUP_locMain" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script language ="javascript" src="../js/validation.js"></script>
<script language ="javascript" src="../js/JS_Calendar.js"></script>


<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Location Definition</title>
    <link rel="stylesheet" href="../stylesheet/newtext.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/button.css" type="text/css">
    <link rel="stylesheet" href="../stylesheet/general.css" type="text/css">

    <script language="javascript" src="../js/formatUtil.js"></script>

    <div runat="server" id="DIVSCRIPT">
     <script type="text/javascript">
         function maskKey(objEvent) {
             var iKeyCode;
             iKeyCode = objEvent.keyCode;

             if ((iKeyCode >= 48 && iKeyCode <= 57)) {
                 return true;
             } else {
                return false;

             }
        }

        function maskKey2(objEvent) {
            var iKeyCode;
            iKeyCode = objEvent.keyCode;

            if ((iKeyCode >= 48 && iKeyCode <= 57) || iKeyCode == 45 || iKeyCode == 46 ) {
                return true;
            } else {
                return false;
            }
        }
    </script>
    </div>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <br />
    <form method="post" name="searchform" id="searchform" onsubmit="" runat="server" defaultbutton="Submit">
    
    <div id="locDiv" runat="server"></div>
    <br />
    <br />
    <br />
    <!--
        **********************
        Modify Here
        -->

    <!--
    **********************
    -->
        
       <table border="0" cellspacing="1" cellpadding="1" align="center" width="70%" id="MainTbl" runat="server" visible="false">
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
                    <asp:Button ID="saveBtn2" runat="server" Text="Save" CssClass="all_button" />
                    <input id="btnBack2" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.close();"
                        class="all_button" /></td>
            </tr>    
            <tr id="flTr" runat="server" visible="false">
                <td>
                        <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
                        <tr>
                            <td class="LabelTD" width="160px" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_NUM" runat="server" /> </font>
                            </td>
                            <td colspan="7">
                                <asp:textbox ID="mFL_NUM" runat="server" Text="" Font-Size="10" MaxLength="3" width="30px" ></asp:textbox>     <%--onkeypress="return maskKey(event);"--%>                           
                            </td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_NAME" runat="server" /> </font>
                            </td>
                            <td colspan="7">
                                <asp:TextBox ID="FL_NAME" runat="server" Width="400" MaxLength="100" />
                            </td>
                        </tr>
                        <tr>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_NAME_CH" runat="server" /> </font>
                            </td>
                            <td colspan="7">
                                <asp:TextBox ID="FL_NAME_CH" runat="server" Width="400" MaxLength="100" />
                            </td>
                        </tr>            
                        <tr>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_GROSS_AREA" runat="server" /></font>
                            </td>
                            <td nowrap>
                                <asp:TextBox ID="FL_GROSS_AREA" runat="server" Width="90px" onkeypress="return maskKey2(event);" />m<sup>2</sup>&nbsp;
                            </td>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_GROSS_CBM" runat="server" /></font>
                            </td>
                            <td colspan="5">
                                <asp:TextBox ID="FL_GROSS_CBM" runat="server" Width="90px" onkeypress="return maskKey2(event);"  />
                            </td>                
                        </tr>     
                         <tr>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_MAX_WEIGHT" runat="server" /> </font>
                            </td>
                            <td colspan="7">
                                <asp:TextBox ID="FL_MAX_WEIGHT" runat="server" Width="100" onkeypress="return maskKey2(event);" />
                            </td>
                        </tr>
                         <tr>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_LENGTH" runat="server" /></font>
                            </td>
                            <td nowrap>
                                <font size="2">
                                    <asp:TextBox ID="FL_LENGTH" runat="server" MaxLength="20" onkeypress="return maskKey2(event);"></asp:TextBox>
                                </font>
                            </td>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_WIDTH" runat="server" /></font>
                            </td>
                            <td nowrap>
                                <font size="2">
                                    <asp:TextBox ID="FL_WIDTH" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                                </font>
                            </td>
                            <td class="LabelTD" nowrap>
                                <font size="2">
                                    <asp:Label ID="lbl_FL_HEIGHT" runat="server" /></font>
                            </td>
                            <td nowrap>
                                <font size="2">
                                    <asp:TextBox ID="FL_HEIGHT" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                                </font>
                            </td>
                        </tr>        
                          <tr> 
                          <td class="LabelTD" nowrap width="160px"><asp:Label ID="lbl_img" runat="server" Text="Image:" /></td>
                          <td colspan="6">
                            <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                  <td>
                                    <br />
                                    <asp:FileUpload ID="FL_PICTURE_upload" runat="server" Width="400px" Font-Size="12px" />
                                      <br />
                                    <asp:Literal ID="FL_PICTURE_lit" runat="server" Visible="false"></asp:Literal>
                                  </td>
                                </tr>
                                 <tr>
                                    <td>
                                        <asp:LinkButton ID="FL_PICTURE_edit" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                                    </td>
                                </tr>
                                 <tr runat="server" id="fl_pic_tr" visible="false">
                                  <td>
                                    <asp:Label ID="fl_preview_pic" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="300" Visible="false" />
                                  </td>
                                </tr>
                                <tr>
                                  <td>&nbsp;</td>
                                </tr>
                                <tr runat="server" id="fl_remove_tr" visible="false">
                                  <td style="padding-bottom: 1px;">
                                    <asp:LinkButton ID="FL_PICTURE_remove" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                                  </td>
                                  </tr>
                                <tr>
                                  <td valign="top">
                                    <asp:HyperLink ID="FL_PICTURE" runat="server" />
                                  </td>
                                </tr>
                           </table>
                          </td>
                          </tr>        
                          </table>
                </td>
            </tr>
            
            <tr id="arTr" runat="server" visible="false">
                <td>
                    <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
                    <tr>
                        <td class="LabelTD" width="160px" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_CODE" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:textbox ID="mAR_CODE" runat="server" Text="" Font-Size="10" MaxLength="3" 
                                Width="30px"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_NAME" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:TextBox ID="AR_NAME" runat="server" Width="400" MaxLength="100" />
                        </td>
                    </tr>
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_NAME_CH" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:TextBox ID="AR_NAME_CH" runat="server" Width="400" MaxLength="100" />
                        </td>
                    </tr>   
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_GROSS_AREA" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <asp:TextBox ID="AR_GROSS_AREA" runat="server" Width="90px" onkeypress="return maskKey2(event);" />m<sup>2</sup>&nbsp;
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_GROSS_CBM" runat="server" /></font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="AR_GROSS_CBM" runat="server" Width="90px" onkeypress="return maskKey2(event);" />
                        </td>                
                    </tr>     
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_X" runat="server" onkeypress="return maskKey(event);" /></font>
                        </td>
                        <td nowrap>
                            <asp:TextBox ID="AR_X" runat="server" Width="90px" /></b>
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_Y" runat="server" /></font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="AR_Y" runat="server" Width="90px" onkeypress="return maskKey(event);" />
                        </td>                
                    </tr>          
                     <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_LENGTH" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="AR_LENGTH" runat="server" MaxLength="20" onkeypress="return maskKey2(event);"></asp:TextBox>
                            </font>
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_WIDTH" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="AR_WIDTH" runat="server" MaxLength="50"></asp:TextBox>
                            </font>
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_HEIGHT" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="AR_HEIGHT" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                            </font>
                        </td>
                    </tr> 
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_AR_FACILITIES" runat="server" /> </font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="AR_FACILITIES" runat="server" Width="400" MaxLength="100" />
                        </td>
                    </tr>    
                    <tr>
                        <%--<td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label runat="server" ID="lbl_AR_DAMAGE_YN" Text="Damage Area" />
                            </font>
                        </td>
                        <td>
                            <asp:CheckBox runat="server" ID="AR_DAMAGE_YN" />
                        </td>--%>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label runat="server" ID="lbl_AR_POWER" Text="Power Y/N" />
                            </font>
                        </td>
                        <td colspan="2" >
                            <asp:CheckBox runat="server" ID="AR_POWER" />
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label runat="server" ID="lbl_AR_SECURITY" Text="Security Level" />
                            </font>
                        </td>
                        <td colspan="2">
                            <asp:DropDownList runat="server" ID="AR_SECURITY" />
                        </td>
                    </tr>     
                    <tr>
                        <%--<td class="LabelTD" nowrap>
                             <font size="2">
                                <asp:Label ID="lbl_AR_TEMP_YN" runat="server" Text="Temperature Control" />:
                             </font>
                        </td>
                        <td>
                            <asp:CheckBox runat="server" ID="AR_TEMP_YN" />
                        </td>--%>
                        <td class="LabelTD" nowrap>
                             <font size="2">
                                <asp:Label ID="lbl_AR_TEMP_FR" runat="server" Text="Temperature From" />:
                             </font>
                        </td>
                        <td colspan="2">
                            <asp:TextBox runat="server" ID="AR_TEMP_FR" MaxLength="20" onkeypress="return maskKey2(event);" />
                        </td>
                        <td class="LabelTD" nowrap>
                             <font size="2">
                                <asp:Label ID="lbl_AR_TEMP_TO" runat="server" text="Temperature To" />:
                             </font>
                        </td>
                        <td colspan="2">
                            <asp:TextBox runat="server" ID="AR_TEMP_TO" MaxLength="20" onkeypress="return maskKey2(event);" />
                        </td>
                    </tr>   
                    <tr>
                        <td class="LabelTD"></td>
                        <td colspan="5">
                            <asp:CheckBox runat="server" ID="AR_SCRAP_AREA" Text="Scrap Area" /><br />
                            <%--<asp:CheckBox runat="server" ID="AR_SHORTLEN_CABLE_AREA" Text="Short Length Cable Area" /><br />--%>
                            <asp:CheckBox runat="server" ID="AR_INSP_AREA" Text="Inspection Area" /><br />
                            <asp:CheckBox runat="server" ID="AR_REPAIR_AREA" Text="For Repair Area" /><br />
                            <%--<asp:CheckBox runat="server" ID="AR_PICKDROP_AREA" Text="For IST Pickup Area" /><br />--%>
                            <asp:CheckBox runat="server" ID="AR_FFI" Text="FFI Area" Checked="true" />                            
                        </td>
                    </tr>
                    <tr>
                       <td class="LabelTD" nowrap width="160px"><asp:Label ID="Label9" runat="server" Text="Image:" /></td>
                       <td colspan="6">
                        <table border="0" cellspacing="0" cellpadding="0" width="100%">
                            <tr>
                              <td>
                                <br />
                                <asp:FileUpload ID="AR_PICTURE_UPLOAD" runat="server" Width="400px" Font-Size="12px" />
                                  <br />
                                <asp:Literal ID="AR_PICTURE_LIT" runat="server" Visible="false"></asp:Literal>
                              </td>
                            </tr>
                             <tr>
                                <td>
                                    <asp:LinkButton ID="AR_PICTURE_EDIT" runat="server" Width="150px" Visible="false">Change Upload Image</asp:LinkButton>
                                </td>
                            </tr>
                             <tr runat="server" id="ar_pic_tr" visible="false">
                              <td>
                                <asp:Label ID="ar_preview_pic" Font-Size="12px" Font-Italic="True" runat="server" Text="Preview: (Click to view actual size)" Width="300" Visible="false" />
                              </td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr runat="server" id="ar_remove_tr" visible="false">
                              <td style="padding-bottom: 1px;">
                                <asp:LinkButton ID="AR_PICTURE_REMOVE" runat="server" Width="150px" Visible="false" OnClientClick="return confirm(&quot;Are you sure to Remove this image?&quot;);">Remove Image</asp:LinkButton>
                              </td>
                              </tr>
                            <tr>
                              <td valign="top">
                                <asp:HyperLink ID="AR_PICTURE" runat="server" />
                              </td>
                            </tr>
                       </table>
                      </td>
                      </tr>   
                      </table>     
                </td>
            </tr>
            
            <tr id="rkTr" runat="server" visible="false">
                <td>
                <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
                    <tr>
                        <td class="LabelTD" width="160px" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_CODE" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:textbox ID="mRK_CODE" runat="server" Text="" Font-Size="10" MaxLength="4" 
                                Width="45px"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_NAME" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:TextBox ID="RK_NAME" runat="server" Width="400" MaxLength="100" />
                        </td>
                    </tr>
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_NAME_CH" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:TextBox ID="RK_NAME_CH" runat="server" Width="400" MaxLength="100" />
                        </td>
                    </tr>   
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_GROSS_AREA" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <asp:TextBox ID="RK_GROSS_AREA" runat="server" Width="90px" onkeypress="return maskKey2(event);"/>m<sup>2</sup>&nbsp;
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_GROSS_CBM" runat="server" /></font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="RK_GROSS_CBM" runat="server" Width="90px" onkeypress="return maskKey2(event);"/>
                        </td>                
                    </tr>     
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_X" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <asp:TextBox ID="RK_X" runat="server" Width="90px" onkeypress="return maskKey(event);" /></td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_Y" runat="server" /></font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="RK_Y" runat="server" Width="90px" onkeypress="return maskKey(event);" />
                        </td>                
                    </tr>          
                     <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_LENGTH" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="RK_LENGTH" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                            </font>
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_WIDTH" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="RK_WIDTH" runat="server" MaxLength="20" onkeypress="return maskKey2(event);"></asp:TextBox>
                            </font>
                        </td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_DEPTH" runat="server" /></font>
                        </td>
                        <td nowrap>
                            <font size="2">
                                <asp:TextBox ID="RK_DEPTH" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                            </font>
                        </td>
                    </tr> 
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_XBINS" runat="server"/></font>
                        </td>
                        <td nowrap>
                            <asp:TextBox ID="RK_XBINS" runat="server" Width="90px"  onkeypress="return maskKey(event);"  /></td>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_YBINS" runat="server"  /></font>
                        </td>
                        <td colspan="5">
                            <asp:TextBox ID="RK_YBINS" runat="server" Width="90px" onkeypress="return maskKey(event);" />
                        </td>                
                    </tr>     
                    <tr>
                        <td class="LabelTD" nowrap>
                            <font size="2">
                                <asp:Label ID="lbl_RK_REM" runat="server" /> </font>
                        </td>
                        <td colspan="7">
                            <asp:TextBox ID="RK_REM" runat="server" Width="400px" TextMode="MultiLine" 
                                MaxLength="100" Height="72px" />
                        </td>
                    </tr>            
                    </table>
                </td>
            </tr>
            
         <tr id="bnTr" runat="server" visible="false">
         <td>
                <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%">
                <tr>
                    <td class="LabelTD" width="160px" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_CODE" runat="server" /> </font>
                    </td>
                    <td>
                        <asp:textbox ID="mBN_CODE" runat="server" Text="" Font-Size="10" MaxLength="3" Width="30px"></asp:textbox>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_STATUS" runat="server" /></font>
                    </td>
                    <td colspan="5">
                        <asp:DropDownList runat="server" ID="BN_STATUS" />
                    </td>                
                </tr>
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_X" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <asp:Label ID="BN_X" runat="server" Width="90px" />
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_Y" runat="server" /></font>
                    </td>
                    <td colspan="5">
                        <asp:Label ID="BN_Y" runat="server" Width="90px" />
                    </td>                
                </tr>     
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_CBM" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <font size="2">
                            <asp:TextBox ID="BN_CBM" runat="server" Width="90px" onkeypress="return maskKey2(event);"/>
                        </font>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_CSMS_CODE" runat="server" /></font>
                    </td>
                    <td nowrap colspan="5">
                        <font size="2">
                            <asp:TextBox ID="BN_CSMS_CODE" runat="server" Width="90px"/>
                        </font>
                    </td>
                </tr> 
                 <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_LENGTH" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <font size="2">
                            <asp:TextBox ID="BN_LENGTH" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                        </font>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_WIDTH" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <font size="2">
                            <asp:TextBox ID="BN_WIDTH" runat="server" MaxLength="20" onkeypress="return maskKey2(event);"></asp:TextBox>
                        </font>
                    </td>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_DEPTH" runat="server" /></font>
                    </td>
                    <td nowrap>
                        <font size="2">
                            <asp:TextBox ID="BN_DEPTH" runat="server" MaxLength="50" onkeypress="return maskKey2(event);"></asp:TextBox>
                        </font>
                    </td>
                </tr> 
                <tr>
                    <td class="LabelTD" nowrap>
                        <font size="2">
                            <asp:Label ID="lbl_BN_REM" runat="server" /> </font>
                    </td>
                    <td colspan="7">
                        <asp:TextBox ID="BN_REM" runat="server" Width="400px" TextMode="MultiLine" 
                            MaxLength="100" Height="72px" />
                    </td>
                </tr>       
                <tr>
                    <td class="LabelTD" nowrap>
                        <asp:Label runat="server" ID="lbl_BN_UTILIZATION_TYPE" text="Utilization Type" />
                    </td>
                    <td colspan="7">
                        <asp:DropDownList runat="server" id="BN_UTILIZATION_TYPE">
                            <asp:ListItem Text="By CBM" Value="CBM" />
                            <asp:ListItem text="By Area" Value="AREA" />
                        </asp:DropDownList>
                    </td>
                </tr>     
                </table>
            </td>
         </tr>  
          
            <tr>
                <td align="left" width="100%">
                    <table border="1" cellspacing="1" cellpadding="0" width="100%">
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
                <td colspan="8" class="menuTD">
                    <asp:Button ID="saveBtn1" runat="server" Text="Save" CssClass="all_button" />
                    <input id="btnBack1" type="button" <%if Session("gLang") = "E" Then %>value="Close"
                        <% Elseif Session("gLang") = "C" Then %>value="返回" <% End If%> onclick="Javascript:window.close();"
                        class="all_button" />
                </td>
            </tr>
        </table>

            <asp:Label ID="testlabel" runat="server" >
        </asp:Label>
     <asp:Button ID="Submit" runat="server" Text="Select" Visible = false class="all_button" UseSubmitBehavior = false />
        <asp:HiddenField ID="pForm" runat ="server" />
        <asp:HiddenField ID="pItemList" runat ="server" />
        <asp:HiddenField ID="wh" runat ="server" />
        <asp:HiddenField ID="WH_CODE" runat ="server" />
        <asp:HiddenField ID="FL_NUM" runat ="server" />
        <asp:HiddenField ID="AR_CODE" runat ="server" />
        <asp:HiddenField ID="RK_CODE" runat ="server" />
        <asp:HiddenField ID="BN_CODE" runat ="server" />
        <asp:HiddenField ID="mode" runat ="server" />
    </form>
</body>
</html>