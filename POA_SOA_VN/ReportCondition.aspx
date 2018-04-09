<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ReportCondition.aspx.cs" Inherits="POA_SOA.ReportCondition" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" >

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
        .style1{width: 100%;
            text-align: left;
        }
        .style2{width: 138px;
            text-align: right;
        }
        .ctl00_Menu1_0 { background-color:white;visibility:hidden;display:none;position:absolute;left:0px;top:0px; }
	    .ctl00_Menu1_1 { color:White;font-family:Verdana;font-size:Medium;text-decoration:none; }
	    .ctl00_Menu1_2 { color:White;background-color:#000099;border-style:None;font-family:Verdana;font-size:Medium;text-decoration:none;width:100%; }
	    .ctl00_Menu1_3 {  }
	    .ctl00_Menu1_4 { padding:2px 5px 2px 5px; }
	    .ctl00_Menu1_5 { width:100%; }
	    .ctl00_Menu1_6 { font-size:Medium; }
	    .ctl00_Menu1_7 { background-color:#000099;padding:0px 5px 0px 5px; }
	    .ctl00_Menu1_8 { background-color:#000099; }
	    .ctl00_Menu1_9 {            text-align: center;
        }
	    .ctl00_Menu1_10 { background-color:Blue; }
	    .ctl00_Menu1_11 {  }
	    .ctl00_Menu1_12 { background-color:#3333FF; }
	    .ctl00_Menu1_13 { color:Blue; }
	    .ctl00_Menu1_14 { color:Blue;background-color:Yellow; }
	    .ctl00_Menu1_15 { color:Blue; }
	    .ctl00_Menu1_16 { color:Blue;background-color:Yellow; }
        .style20
        {
            width: 130px;
            text-align: right;
        }
        .style21
        {
            width: 414px;
        }
        .style22
        {
            width: 138px;
            text-align: right;
            height: 30px;
        }
        .style23
        {
            height: 30px;
        }
        .style24
        {
            width: 138px;
            text-align: right;
            height: 22px;
        }
        .style25
        {
            height: 22px;
        }
        .style26
        {
            width: 305px;
        }
        .style29
        {
            width: 138px;
            text-align: right;
            height: 23px;
        }
        .style30
        {
            height: 23px;
        }
        .style32
        {
            height: 22px;
            width: 414px;
        }
        .style33
        {
            color: #0000FF;
            font-style: italic;
        }
        .style34
        {
        }
        </style>
    </head>
<body>
    <form id="form1" runat="server" style="width:100%">
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
    
        <br />	
    <table class="style1"> <tr>
    <td colspan="2">
       <table id="ctl00_Menu1" class="ctl00_Menu1_5 ctl00_Menu1_2" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td onmouseover="Menu_HoverRoot(this)" onmouseout="Menu_Unhover(this)" onkeyup="Menu_Key(event)" title="Home" id="ctl00_Menu1n0"><table class="ctl00_Menu1_4 ctl00_Menu1_10" cellpadding="0" cellspacing="0" border="0" width="100%">
			<tr>
				<td style="white-space:nowrap;"><a class="ctl00_Menu1_1 ctl00_Menu1_3 ctl00_Menu1_9" href="/VN/DMS/Default.aspx">Home</a></td>
			</tr>
		</table></td><td style="width:3px;"></td><td><table border="0" cellpadding="0" cellspacing="0" width="100%" class="ctl00_Menu1_5">
			<tr>
				<td onmouseover="Menu_HoverStatic(this)" onmouseout="Menu_Unhover(this)" onkeyup="Menu_Key(event)" title="Delivery Management" id="ctl00_Menu1n1"><table class="ctl00_Menu1_4" cellpadding="0" cellspacing="0" border="0" width="100%">
					<tr>
						<td style="white-space:nowrap;"><a class="ctl00_Menu1_1 ctl00_Menu1_3" href="/VN/DMS/JDE/Default.aspx" style="margin-left:5px;">Delivery Management</a></td>
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
                        <span style="color:#3366FF;font-weight:bold;">Select order information</span><a id="ctl00_SiteMapPath1_SkipLink"></a></span>
                    </td></tr>

    </table>
    <table class="style1">
        <tr>
            <td colspan="5"> Please Fill-in conditions for your request:<br /></td>
        </tr>
        
            <tr>
                <td class="style2">JDE Order Number:</td>
                <td class="style21">
                    <asp:TextBox ID="txtOrderNo" runat="server"></asp:TextBox>&nbsp;&nbsp; <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" ControlToValidate="txtOrderNo"
                            ErrorMessage="*Enter number only" ValidationExpression="^\d*" ForeColor="Red" 
                        style="color: #000099"></asp:RegularExpressionValidator>                    
                    
                    </td>
                <td class="style20">
                    Order Type:
                    </td>
                <td class="style26">
                    <asp:DropDownList AutoPostBack="true" ID="dlistOrderType" runat="server" >
                        <asp:ListItem>SD</asp:ListItem>
                        <asp:ListItem>SX</asp:ListItem>
                        <asp:ListItem>OD</asp:ListItem>
                        <asp:ListItem>OX</asp:ListItem>
                        <asp:ListItem>SB</asp:ListItem>
                        <asp:ListItem>OB</asp:ListItem>
                        <asp:ListItem>SO</asp:ListItem>
                        <asp:ListItem>OP</asp:ListItem>
                    </asp:DropDownList>
                    </td>
            </tr>
            <tr>
                <td class="style22">
        Account Payable:</td>
                <td class="style23">
                    <asp:DropDownList ID="dlistAccount" runat="server" 
                        DataSourceID="SqlData_Account" DataTextField="Account_Name" 
                        DataValueField="Account_Code" Height="26px" Width="392px">
                    </asp:DropDownList>
                    </td>
                <td class="style20">
                    Term:</td>
                <td class="style23" colspan="2">
                    <asp:DropDownList ID="dlistTerm" runat="server" Height="26px" Width="351px" 
                        DataSourceID="SqlData_FormTerm" DataTextField="FieldContent" 
                        DataValueField="FieldContent">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style22">
                    <asp:Button ID="btnValid" runat="server" Text="Check Validity" 
                        onclick="btnValid_Click" />
                </td>
                <td class="style23">
                    <asp:Label ID="lblValidData" runat="server" style="color: #FF0000"></asp:Label>
                    
                    
                    </td>
                <td class="style20">
                    Address:</td>
                <td class="style23" colspan="2">
                    <asp:DropDownList ID="dlistBranch" runat="server" 
                        DataSourceID="SqlData_Address" DataTextField="Name" DataValueField="Code" 
                        Height="26px" Width="350px" 
                        onselectedindexchanged="dlistBranch_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style2">
                    &nbsp;</td>
                <td class="style21">
                    <asp:CheckBox ID="chkUSDB50" runat="server" 
                        style="text-align: right"
                        oncheckedchanged="chkUSDB50_CheckedChanged" AutoPostBack="True" 
                        Text="B50: Print USD only" />
                </td>
                <td class="style20">
                    Exchange Rate:</td>
                <td class="style26"> 
                    <asp:TextBox ID="txtExRate" runat="server" Width="110px" AutoPostBack="True" 
                        Enabled="False" ontextchanged="txtExRate_TextChanged" ></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                    ControlToValidate="txtExRate"  ErrorMessage="Exchange Rate is required." 
                        style="color: #0000CC" Enabled="False"> *Please input </asp:RequiredFieldValidator>
                    
                    <asp:RegularExpressionValidator 
            ID="RegularExpressionValidator3" runat="server" ControlToValidate="txtExRate" Type="Currency"
        ErrorMessage="*Number only" ValidationExpression="^\d+(\.\d{1,2})?$" 
                        style="color: #000099"></asp:RegularExpressionValidator>
                    </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    &nbsp;</td>
                <td class="style21">
                    <asp:CheckBox ID="chkForeignCur" runat="server" 
                        style="text-align: right" AutoPostBack="True" 
                        Text="Print foreign currency only" Visible="False" />
                    </td>
                <td class="style20">
                    <asp:CheckBox ID="chkGrpLotNo" runat="server" AutoPostBack="True" 
                        oncheckedchanged="chkGrpLotNo_CheckedChanged" 
                        style="font-weight: 700; text-align: right" Text="JDE Lot No." />
                    </td>
                <td class="style26"> 
                    <asp:RadioButtonList ID="rdlsNoOfLot" runat="server" Enabled="False" 
                        Height="16px" RepeatDirection="Horizontal" style="font-style: italic" 
                        Width="381px">
                        <asp:ListItem Selected="True" Value="1">Lot Number</asp:ListItem>
                        <asp:ListItem Value="2">Lot No. 1 - Lot No. 2</asp:ListItem>
                    </asp:RadioButtonList>
                    </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    Tax Code:</td>
                <td class="style21">
                    <asp:TextBox ID="txtTaxable" runat="server" ReadOnly="True" Width="34px"></asp:TextBox>
                    <asp:TextBox ID="txtTaxCode" runat="server" ReadOnly="True" Width="70px"></asp:TextBox>
                    &nbsp; Tax Rate:<asp:DropDownList ID="dlistTax" runat="server" DataSourceID="SqlData_TaxRate" 
                        DataTextField="TaxRate" DataValueField="JDETaxCode" Enabled="False" 
                        onselectedindexchanged="dlistTax_SelectedIndexChanged" Width="87px">
                    </asp:DropDownList>
                    
                    
                    %</td>
                <td class="style20">
                    Currency:</td>
                <td class="style26"> 
                    <asp:TextBox ID="txtCurrency" runat="server" ReadOnly="True" Width="78px" 
                        style="text-align: left"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    JDE CS Rep:</td>
                <td class="style21">
                    <asp:TextBox ID="txtCSRep" runat="server" ReadOnly="True"></asp:TextBox>
                </td>
                <td class="style20">
                    Branch:</td>
                <td class="style34">
                    <asp:TextBox ID="txtBranch" runat="server" ReadOnly="True" Width="89px" 
                        style="text-align: left"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    Type of document:</td>
                <td class="style21">
                    <asp:Label ID="lblPartner" runat="server"></asp:Label>
                    <asp:Label ID="lblPartnerCode" runat="server"></asp:Label>
                </td>
                <td class="style20">
                    &nbsp;Customer 
                    PO:</td>
                <td class="style26">
                    <asp:TextBox ID="txtCustomerPO" runat="server" ReadOnly="True" Width="280px" 
                        style="text-align: left"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2">
                    </td>
                <td class="style21">
                    <asp:DropDownList ID="dListGroupName" runat="server" 
                        DataSourceID="SqlData_DetailMail" DataTextField="GroupName" 
                        DataValueField="GroupName" Enabled="False" Width="392px" >
                    </asp:DropDownList>
                </td>
                <td class="style20">
                    Sending List:</td>
                <td class="style26">
                    <asp:DropDownList ID="dlistCCList" runat="server" Height="25px" Width="280px" 
                        DataSourceID="SqlData_MailList" DataTextField="Description" 
                        DataValueField="Description" 
                        onselectedindexchanged="dlistCCList_SelectedIndexChanged" OnDataBound="dlistCCList_DataBound"
                        AutoPostBack="True">
                        <asp:ListItem>-- Select a list to CC --</asp:ListItem>
                    </asp:DropDownList>

                    </td>
                <td class="style33">
                    *Select the list</td>
            </tr>
            <tr>
                <td class="style24">
                    To: </td>
                <td class="style32">
                    <asp:ListBox ID="lstToSelect" runat="server" Height="110px" Width="260px" 
                        Enabled="False"></asp:ListBox>
                    <br />
                    <asp:Label ID="lblSenderEmail" runat="server"></asp:Label>
                </td>
                <td class="style20">
                    CC: </td>
                <td class="style26">
                    <asp:ListBox ID="lstCCSelect" runat="server" Height="110px" Width="280px" 
                        Enabled="False" DataSourceID="SqlData_DetailMail" DataTextField="CCList" 
                        DataValueField="CCList" ></asp:ListBox>
                </td>
                <td class="style25">
                    <asp:Button ID="btnClearCC" runat="server" Text="Clear All" 
                        onclick="btnClearCC_Click" Visible="False" />
                    <asp:Button ID="btnAddCCList" runat="server" Text="Add" 
                        onclick="btnAddCCList_Click" Visible="False" />
                    </td>
            </tr>
            <tr>
                <td class="style29">
                    </td>
                <td colspan="4" class="style30">
                    
                    
        <asp:Button ID="btnPrint" runat="server" 
            Text="Get Form" onclick="btnPrint_Click"
            OnClientClick="aspnetForm.target ='_blank';" Width="75px" style="text-align: right" 
                        Enabled="False" />
                    
                    
                    </td>
            </tr>
            <tr>
                <td class="style2">
                    &nbsp;</td>
                <td colspan="4"></td>
            </tr>
            <tr>
                <td class="style2">
                    
                    
                    &nbsp;</td>
                <td colspan="4">
                    &nbsp;</td>
            </tr>
        </table>
        &nbsp;<br />
        <asp:SqlDataSource ID="SqlData_Account" runat="server" 
            ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>"          
            
            
            
            
            
            
        SelectCommand="SELECT Account_Code, Account_Short_Name, Account_Name, Address, Address_Line2, Bank_Name, Bank_Address, Account_USD, Account_LocalCur, Swift_Code, Account_Country FROM Account WHERE (UPPER(Account_Country) IN ('ALL', 'VN','INDIA','INDO'))"></asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlData_MailList" runat="server" 
            ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>" 
            
            SelectCommand="SELECT * FROM [SendingList] WHERE ([JDECode] = @JDECode)">
            <FilterParameters>
                <asp:ControlParameter ControlID="lblPartnerCode" DefaultValue="&quot;&quot;" 
                    Name="JDEAddCode" PropertyName="Text" />
            </FilterParameters>
            <SelectParameters>
                <asp:ControlParameter ControlID="lblPartnerCode" Name="JDECode" 
                    PropertyName="Text" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlData_TaxRate" runat="server" 
            ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>" 
            
            
            SelectCommand="SELECT CONVERT (VARCHAR(10), 100 * TaxRate) + '%' + ' : ' + JDETaxCode AS Tax, JDETaxCode , Branch, Location, TaxRate, TypeOfObject, Note, CustomerTax, SupplierTax FROM TaxRate ORDER BY TaxRate">
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlData_FormTerm" runat="server" 
        ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>" 
        SelectCommand="SELECT FieldName, FieldContent, Country, DocType, Branch, Note FROM FormPrinting WHERE (FieldName = 'Term')">
    </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlData_DetailMail" runat="server" 
            ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>" 
            
            SelectCommand="SELECT JDECode, GroupName, TOList, CCList, ReportType, Country, USCode, UserDefCode, Description, ModifiedTime, ModifiedBy FROM SendingList WHERE (JDECode = @JDECode) AND (Description = @Description)">
            <SelectParameters>
                <asp:ControlParameter ControlID="lblPartnerCode" Name="JDECode" 
                    PropertyName="Text" Type="String" />
                <asp:ControlParameter ControlID="dlistCCList" Name="Description" 
                    PropertyName="SelectedValue" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="SqlData_Address" runat="server" 
            ConnectionString="<%$ ConnectionStrings:OPD_DBConnectionString1 %>" 
            
            SelectCommand="SELECT [Code], [Name] FROM [Address] WHERE ([Location] = @Location)">
            <SelectParameters>
                <asp:ControlParameter ControlID="txtBranch" DefaultValue="B40" Name="Location" 
                    PropertyName="Text" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <br />
    
    
    <table style="text-align: center;width:100%">   
             <tr><td>        
                    @ 2017 L&E International Ltd. All rights reserved.
                    &nbsp;</td>
            </tr> 
     </table>
    </form>
</body>
</html>
