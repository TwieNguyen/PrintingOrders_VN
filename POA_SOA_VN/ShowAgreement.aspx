<%@ Page ClientTarget="uplevel" Language="C#" AutoEventWireup="true" CodeBehind="ShowAgreement.aspx.cs" Inherits="POA_SOA.ShowAgreement" %>
<%@ Register assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692FBEA5521E1304" namespace="CrystalDecisions.Web" tagprefix="CR" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<script type="text/javascript">
//<![CDATA[
    var theForm = document.forms['aspnetForm'];
    if (!theForm) {
        theForm = document.aspnetForm;
    }
    function __doPostBack(eventTarget, eventArgument) {
        if (!theForm.onsubmit || (theForm.onsubmit() != false)) {
            theForm.__EVENTTARGET.value = eventTarget;
            theForm.__EVENTARGUMENT.value = eventArgument;
            theForm.submit();
        }
    }
//]]>
</script>

<head runat="server">
    <title></title>
    <style type="text/css">
	    .ctl00_Menu1_0 { background-color:white;visibility:hidden;display:none;position:absolute;left:0px;top:0px; }
	    .ctl00_Menu1_1 { color:White;font-family:Verdana;font-size:Medium;text-decoration:none; }
	    .ctl00_Menu1_2 { color:White;background-color:#000099;border-style:None;font-family:Verdana;font-size:Medium;text-decoration:none;width:100%; }
	    .ctl00_Menu1_3 {  }
	    .ctl00_Menu1_4 { padding:2px 5px 2px 5px; }
	    .ctl00_Menu1_5 { width:100%; }
	    .ctl00_Menu1_6 { font-size:Medium; }
	    .ctl00_Menu1_7 { background-color:#000099;padding:0px 5px 0px 5px; }
	    .ctl00_Menu1_8 { background-color:#000099; }
	    .ctl00_Menu1_9 {  }
	    .ctl00_Menu1_10 { background-color:Blue;
            width: 249px;
        }
	    .ctl00_Menu1_11 {  }
	    .ctl00_Menu1_12 { background-color:#3333FF; }
	    .ctl00_Menu1_13 { color:Blue; }
	    .ctl00_Menu1_14 { color:Blue;background-color:Yellow; }
	    .ctl00_Menu1_15 { color:Blue; }
	    .ctl00_Menu1_16 { color:Blue;background-color:Yellow; }
        .style1
        {
            width: 100%;
        }
        .style2
        {
            width: 141px;
        }
        .style3
        {}
        .style4
        {
            width: 250px;
        }
        .style7
        {
            font-weight: bold;
            text-decoration: underline;
        }
        </style>
</head>
<body>
    <table style="width:100%">
            <tr>
                <td style="width: 40%"><asp:Image ID="Image1" runat="server" ImageUrl="~/img/headerlogo.png" href="/VN/DMS/Default.aspx"/></td>
                <td align="right" style="width: 60%" valign="top">
                    <span id="ctl00_lblUserName" style="font-size:Smaller;">Welcome&nbsp;&nbsp;</span><%=Request.QueryString["UserName"]%><br />                    
                    <br /><br />
                    <a id="ctl00_LoginStatus1" href="/VN/DMS/Login.aspx" style="font-size:Smaller;">Logout</a>
                    <a id="ctl00_hplChangePassword" href="/VN/DMS/System/ChangePassword.aspx" style="font-size:Smaller;">Change Password</a>
                    <a id="ctl00_hplContactUs" href="mailto:tom.leong@le-intl.com" style="font-size:Small;">Contact Us</a>                                    
                </td>                                    
            </tr>
    </table>
            
   <table class="style1"> <tr>
    <td colspan="2">
       <table id="ctl00_Menu1" class="ctl00_Menu1_5 ctl00_Menu1_2" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td onmouseover="Menu_HoverRoot(this)" onmouseout="Menu_Unhover(this)" 
            onkeyup="Menu_Key(event)" title="Home" id="ctl00_Menu1n0" class="style4"><table class="ctl00_Menu1_4 ctl00_Menu1_10" cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td style="white-space:nowrap;"><a class="ctl00_Menu1_1 ctl00_Menu1_3 ctl00_Menu1_9" href="/VN/DMS/Default.aspx">Home</a></td>
			</tr>
		</table></td><td style="width:3px;"></td><td><table border="0" cellpadding="0" cellspacing="0" width="100%" class="ctl00_Menu1_5">
			<tr>
				<td onmouseover="Menu_HoverStatic(this)" onmouseout="Menu_Unhover(this)" onkeyup="Menu_Key(event)" title="Delivery Management" id="ctl00_Menu1n1"><table class="ctl00_Menu1_4" cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td style="white-space:nowrap;"><a class="ctl00_Menu1_1 ctl00_Menu1_3" href="/VN/POA_SOA/ReportCondition.aspx?UserName=<%=Request.Cookies["Info"]["Username"].ToString().Trim()%>" style="margin-left:5px;">Send/Print More POA-SOA</a></td>
					</tr>
				</table></td><td style="width:3px;"></td><td onmouseover="Menu_HoverStatic(this)" onmouseout="Menu_Unhover(this)" onkeyup="Menu_Key(event)" title="Account Management" id="ctl00_Menu1n2"><table class="ctl00_Menu1_4" cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td style="white-space:nowrap;"><a class="ctl00_Menu1_1 ctl00_Menu1_3" href="/VN/DMS/Default.aspx" style="margin-left:5px;">Account Management</a></td>
					</tr>
				</table></td>
			</tr>
		</table></td>
	</tr>
</table><a id="ctl00_Menu1_SkipLink"></a>
                        <span id="ctl00_SiteMapPath1" style="display:inline-block;background-color:Silver;font-family:Verdana;font-size:Small;width:100%;">
                        <span style="color:#3366FF;font-weight:bold;">Send order information</span><a id="ctl00_SiteMapPath1_SkipLink"></a></span>
                    </td></tr>

    </table>         
   
         
    <form id="form1" runat="server">
    <div>
    
        <table class="style1">
            <tr>
                <td class="style2">
    
                    
                </td>
                <td class="style3">
                    <asp:Button ID="btnGetPDF" runat="server" style="margin-left: 0px" 
                        Text="Get PDF" onclick="btnGetPDF_Click" />
                    &nbsp;&nbsp;
    
        <asp:Button ID="btnExportExcel" runat="server" onclick="btnExportExcel_Click" Text="Get Excel" />
                &nbsp;&nbsp;
                    </td>
            </tr>
            <tr>
                <td class="style2">
                    To:</td>
                <td>
                    <asp:Label ID="lblToList" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    CC: (Testing)</td>
                <td>
                    <asp:Label ID="lblCCList" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    Subject:</td>
                <td>
                    <asp:TextBox ID="txtMailSubject" runat="server" Width="700px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    Note to sender:<br />
                    (No <b><i><u> special</u></i></b> character<br />
                    /\()&quot;~!@#$%&#39;^&amp;*)</td>
                <td>
                    <asp:TextBox ID="txtEmailNote" runat="server" Height="50px" Width="700px"></asp:TextBox>
                    
                    <asp:RegularExpressionValidator 
            ID="RegularExpressionValidator3" runat="server" ControlToValidate="txtEmailNote" Type="Currency"
        ErrorMessage="*Remove special character" ValidationExpression="^[^/\\()&quot;~!@#$%'^&amp;*]*$" 
                        style="color: #000099"></asp:RegularExpressionValidator>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    <asp:Label ID="lblUsername" runat="server"></asp:Label>
                </td>
                <td>
                    <asp:Button ID="btnEmailBody" runat="server" Text="Preview Email"  Width="97px" 
                        onclick="btnEmailBody_Click" />
                </td>
            </tr>
            <tr>
                <td class="style2">
                    &nbsp;</td>
                <td>
                    
                    &nbsp;</td>
            </tr>
            </table>     
           
    <span class="style7">Email Preview:</span> <br /><br /> 
    <asp:Label ID="lblEmailBody" runat="server" Visible="False" />    <br /> <br /> 
    <span class="style7">Attachment content:</span>
    <CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" EnableDatabaseLogonPrompt="False" 
            EnableParameterPrompt="False" ReuseParameterValuesOnRefresh="True" 
            ToolPanelView="None" />                
    <br />
    
    <asp:Button ID="btnSendEmail" runat="server" style="margin-left: 0px" 
            Text="Send Email" onclick="btnSendEmail_Click" Enabled="False" />
    <br /> Email sending status:  <asp:Label ID="lblEmailNotify" runat="server" 
            style="color: #FF0000"></asp:Label>
    <table style="text-align: center;width:100%">   
             <tr><td>        
                    @ 2017 L&E International Ltd. All rights reserved.
                    &nbsp;</td>
            </tr> 
     </table>
     </form> 
</body>

</html>
