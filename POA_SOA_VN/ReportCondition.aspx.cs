using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using Microsoft.VisualBasic;
using POA_SOA.UserCode;

namespace POA_SOA
{
    public partial class ReportCondition : System.Web.UI.Page
    {
        public static string OrderNumber, OrderType, Branch, TaxCode, Currency, JDEAddCode, ToList, CCList, CustomerPO, SenderEmail ;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ClearData();
            }

        }


        protected void btnPrint_Click(object sender, EventArgs e)
        {
            Response.Cookies["Info"]["OrderNumber"] = txtOrderNo.Text.Trim();
            Response.Cookies["Info"]["OrderType"] = dlistOrderType.Text.Trim();
            Response.Cookies["Info"]["OrderNumber"] = txtOrderNo.Text.Trim();
            Response.Cookies["Info"]["OrderType"] = dlistOrderType.Text.Trim();
            Response.Cookies["Info"]["Account"] = dlistAccount.SelectedValue.ToString();
            Response.Cookies["Info"]["BranchAddress"] = dlistBranch.SelectedValue.ToString();
            Response.Cookies["Info"]["Curr"] = txtCurrency.Text.Trim();
            Response.Cookies["Info"]["Taxable"] = txtTaxable.Text.Trim();
            Response.Cookies["Info"]["Term"] = dlistTerm.SelectedValue.ToString();
            Response.Cookies["Info"]["Branch"] = txtBranch.Text.Trim();
            if (chkUSDB50.Checked)
                { Response.Cookies["Info"]["USDOnly"] = "1"; }
            else
                { Response.Cookies["Info"]["USDOnly"] = "0"; }

            Response.Cookies["Info"]["TaxRate"] = dlistTax.SelectedItem.Text;
            Response.Cookies["Info"]["CustPO"] = CustomerPO;
            Response.Cookies["Info"]["UserName"] = Request.QueryString["UserName"];

            if ((txtTaxCode.Text != "") && (!chkUSDB50.Checked)) Response.Cookies["Info"]["ExRate"] = txtExRate.Text.Trim();
            else Response.Cookies["Info"]["ExRate"] = "0";

            if (chkGrpLotNo.Checked)
                Response.Cookies["Info"]["UseLotNo"] = (string) rdlsNoOfLot.SelectedValue;
            else
                Response.Cookies["Info"]["UseLotNo"] = "0";
            //Get To list and CC list
            ToList = "";
            CCList = "";
            for (Int16 i = 0; i < lstToSelect.Items.Count; i++)
            {
                ToList += lstToSelect.Items[i].ToString() + ",";
            }
            ToList += lblSenderEmail.Text;

            for (Int16 i = 0; i < lstCCSelect.Items.Count; i++)
            {
                CCList += lstCCSelect.Items[i].ToString() + ",";
            }

            if ((CCList == "") && ((Branch == "B40") || (Branch == "B50"))) //|| (Branch == "B120")))
            {
                lblValidData.Text = "Please check email addresses in TO and CC before sending out";
            }
            else
            {
                lblValidData.Text = "";
                Response.Cookies["Info"]["ToList"] = ToList;
                Response.Cookies["Info"]["CCList"] = CCList;
                Response.Cookies["Info"].Expires = DateTime.Now.AddMinutes(10);
                Response.Redirect("ShowAgreement.aspx");
            }

            

            /*
            Session["OrderNumber"] = txtOrderNo.Text.Trim();
            Session["OrderType"] = dlistOrderType.Text.Trim();
            Session["Account"] = dlistAccount.SelectedValue.ToString();
            Session["BranchAddress"] = dlistBranch.SelectedValue.ToString();
            Session["Curr"] = txtCurrency.Text.Trim();            
            Session["Taxable"] = txtTaxable.Text.Trim();
            Session["Term"] = dlistTerm.SelectedValue.ToString();
            Session["Branch"] = txtBranch.Text.Trim();
            Session["USDOnly"] = chkUSDB50.Checked;
            Session["TaxRate"] = dlistTax.SelectedItem.Text;
            Session["CustPO"] = CustomerPO;
            Session["Username"] = Request.QueryString["UserName"];

            if ((txtTaxCode.Text != "") && (!chkUSDB50.Checked)) Session["ExRate"] = txtExRate.Text.Trim();
                else Session["ExRate"] = 0;

            if (chkGrpLotNo.Checked)
                Session["UseLotNo"] = rdlsNoOfLot.SelectedValue;
            else
                Session["UseLotNo"] = 0;
            //Get To list and CC list
            ToList = "";
            CCList = "";
            for (Int16 i = 0; i < lstToSelect.Items.Count; i++)
            {
                ToList += lstToSelect.Items[i].ToString() + ",";
            }
            ToList += lblSenderEmail.Text;

            for (Int16 i = 0; i < lstCCSelect.Items.Count; i++)
            {
                CCList += lstCCSelect.Items[i].ToString() + ",";
            }

            if ((CCList == "") && ((Branch=="B40")||(Branch=="B50")))
            {
                lblValidData.Text = "Please check email addresses in TO and CC before sending out";
            }
            else            
            {                
                lblValidData.Text = "";
                Session["ToList"] = ToList;
                Session["CCList"] = CCList;
                Response.Redirect("ShowAgreement.aspx");
            }

            */

                                    
        }

        protected void btnValid_Click(object sender, EventArgs e)
        {
            OrderNumber = txtOrderNo.Text.Trim();
            OrderType = dlistOrderType.Text.Trim();

            ClearData();
            if (isDataValid()) 
            {                
                
                //Get Tax Rate
                if (TaxCode == "") dlistTax.SelectedItem.Text = "0";
                else dlistTax.SelectedValue = TaxCode;

                switch (Branch.ToUpper())
                { 
                    case "B40":                        
                        if (TaxCode.ToUpper() != "")
                        {
                            //lblValidData.Visible = true;
                            lblValidData.Text = "B40 Order with Tax - Error!";
                            btnPrint.Enabled = false;
                        }
                        btnPrint.Enabled = true;
                        break;
                    case "B50":                        
                        if (txtTaxCode.Text != "") 
                        {
                            txtExRate.Enabled = true;
                            RegularExpressionValidator3.Enabled = true;
                            RequiredFieldValidator1.Enabled = true;
                            txtExRate.Focus();
                            btnPrint.Enabled = false;
                        }
                        else
                        {
                            txtExRate.Enabled = false;
                            RegularExpressionValidator3.Enabled = false;
                            RequiredFieldValidator1.Enabled = false;
                            btnPrint.Enabled = true;
                        }
                        break;
                    case "B120":
                        
                        //if (TaxCode == "") dlistTax.SelectedItem.Text = "0";
                            //else dlistTax.SelectedValue = TaxCode;
                        btnPrint.Enabled = true;
                        break;
                    default:
                        //lblValidData.Visible = true;
                        //lblValidData.Text = "Just can print for B40, B50 and B120";
                        btnPrint.Enabled = true;
                        break;
                }
               
                getToList();
                
            }

            else
                
            {
                btnPrint.Enabled = false;
                lblPartner.Visible = false;
            }

            

        }

        public bool isDataValid()
        {
            bool Result = false;
            string Type;
            //Check Data validity
            string tSQL;
            if (Strings.Left(OrderType, 1) == "O")
                Type = "POA";
            else
                Type = "SOA";
            tSQL = "EXEC [dbo].[Proc_SHOW_" + Type + "] '" + OrderNumber + "', '" + OrderType + "'";
            SqlConnection oSqlConnection = null;
            SqlCommand cmd = null;
            try
            {
                oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
                oSqlConnection.Open();
                cmd = new SqlCommand(tSQL, oSqlConnection);
                SqlDataReader oReader = cmd.ExecuteReader();

                //Valid data in database
                if (oReader.HasRows)
                {
                    //Get data to variable
                    if (oReader.Read())
                    {
                        Result = true;
                        Branch = Convert.ToString(oReader["Branch"]).Trim();
                        TaxCode = Convert.ToString(oReader["TaxCode1"]).Trim();
                        Currency = Convert.ToString(oReader["CurrencyCode"]).Trim();
                        //Get Supplier Code or Customer Code
                        if (Type == "POA")
                        {
                            JDEAddCode = Convert.ToString(oReader["SupplierCode"]).Trim();
                            lblPartner.Text = "POA - TO SUPPLIER: ";
                            CustomerPO = Convert.ToString(oReader["CustomerPO"]).Trim();
                        }
                        else
                        {
                            JDEAddCode = Convert.ToString(oReader["ShipToCode"]).Trim();
                            lblPartner.Text = "SOA - TO SHIP TO: ";
                            CustomerPO = Convert.ToString(oReader["Reference1"]).Trim();
                        }
                        txtBranch.Text = Branch;
                        txtTaxCode.Text = TaxCode;
                        txtCurrency.Text = Currency;
                        txtTaxable.Text = Convert.ToString(oReader["Taxable"]).Trim();
                        txtCSRep.Text = Convert.ToString(oReader["JDEUserID"]).Trim();
                        lblPartnerCode.Text = JDEAddCode;
                        txtCustomerPO.Text = CustomerPO;

                    }   
       
                }
                else
                {
                    //empty data
                    Result = false;
                    ClearData();
                    //lblValidData.Visible = true;
                    lblValidData.Text = "Cannot find order data. Please check or wait for data uploading!";
                    
                }
            }
            catch (Exception ex)
            {
                //lblValidData.Text = ex.Message;
            }
            finally
            {
                if (oSqlConnection != null) oSqlConnection.Dispose();
                if (cmd != null) oSqlConnection.Dispose();
            }

            return Result;
        }

        protected void txtExRate_TextChanged(object sender, EventArgs e)
        {
            if (txtExRate.Enabled)
                if (txtExRate.Text != "") btnPrint.Enabled = true;
        }

        void ClearData()
        {
            //lblValidData.Visible = false;
            lblValidData.Text = "";
            dListGroupName.Visible = true;
            btnPrint.Enabled = false;
            txtCSRep.Text = "";
            txtCurrency.Text = "";
            txtTaxable.Text = "";
            txtTaxCode.Text = "";
            txtExRate.Text = "";
            txtBranch.Text = "";
            txtCustomerPO.Text = "";
            lstCCSelect.Items.Clear();
            lstToSelect.Items.Clear();
        
        }

        protected void chkUSDB50_CheckedChanged(object sender, EventArgs e)
        {
            if ((!chkUSDB50.Checked)&&(txtBranch.Text=="B50"))
            {
                txtExRate.Enabled = true; 
                btnPrint.Enabled=false;
            }
            else 
            {
                txtExRate.Enabled = false; 
                btnPrint.Enabled=true;
            }
            
        }

        void getToList()
        { 
            //Get all email of the supplier code into the To list & creator
            SqlConnection oSqlConnection = null;
            string tSQL = "";
            
            string UName = Request.QueryString["UserName"].ToString().Trim().ToUpper();
            //if (UName!="") DMSUserName = UName;
            try
            {                
                //email of creator
                tSQL = "SELECT [FirstName],[LastName],[EmailAddress],[Department] FROM [User_DB].[dbo].[vw_aspnet_LEEmailAddress] " +
                        " WHERE  RTRIM(Upper(FirstName))+'.'+RTRIM(Upper(LastName)) = '" + UName + "'"; //"' AND "+ "RTRIM(Department)IN ('CS','IT')";
                oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
                SqlDataAdapter dap = new SqlDataAdapter(tSQL, oSqlConnection);
                DataTable dt = new DataTable();
                dap.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    //Get data to variable
                    //lstToSelect.Items.Add(dt.Rows[0]["EmailAddress"].ToString());       
                    lblSenderEmail.Text = dt.Rows[0]["EmailAddress"].ToString();
                }

                //Get TO list by Order Type
                if (Strings.Left(OrderType, 1) == "O")
                {
                    //Unbind any datasource
                    lstToSelect.DataSource = null;

                    //Get data to variable
                    //lstToSelect.Items.Add(dt.Rows[0]["EmailAddress"].ToString());

                    tSQL = "SELECT [JDEAddressCode],[FirstName],[LastName],[Company],[EmailAddress],[Location] FROM [User_DB].[dbo].[vw_aspnet_SupplierEmailAddress] " +
                           " WHERE [JDEAddressCode]='" + JDEAddCode.Trim() + "'";
                    oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
                    dap = new SqlDataAdapter(tSQL, oSqlConnection);
                    dt = new DataTable();
                    dap.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        //Get data to variable
                        for (Int16 i = 1; i <= dt.Rows.Count; i++)
                        {
                            lstToSelect.Items.Add(dt.Rows[i - 1]["EmailAddress"].ToString());
                        }
                    }
                }
                else
                //for Sales Order
                {
                    dListGroupName.Visible = true;
                    lstToSelect.DataSourceID = "SqlData_DetailMail";
                    lstToSelect.DataTextField = "TOList";
                    lstToSelect.DataBind();
                }
            }

            catch (Exception ex)
            {
                
                lblValidData.Text = ex.Message;
            }
            finally
            {
                if (oSqlConnection != null) oSqlConnection.Dispose();
            }

        
        }


        protected void dlistTax_SelectedIndexChanged(object sender, EventArgs e)
        {
            


        }

       
        protected void btnAddCCList_Click(object sender, EventArgs e)
        {
            lstCCSelect.Items.Add(dlistCCList.SelectedValue.ToString());
        }

       

        protected void btnClearCC_Click(object sender, EventArgs e)
        {
            lstCCSelect.Items.Clear();
        }

        protected void dlistCCList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dlistCCList.SelectedIndex == 0)//empty Item
            { 
                if (Strings.Left(OrderType, 1) == "S")
                {
                    dListGroupName.Visible = false;
                    lstToSelect.Items.Clear();
                    lstCCSelect.Items.Clear();
                }
                else
                {
                    dListGroupName.Visible = false;
                    lstCCSelect.Items.Clear();
                }
            }           

        }

        protected void dlistCCList_DataBound(object sender, EventArgs e)
        {
            DropDownList ddl = (DropDownList)sender;
            ListItem emptyItem = new ListItem("-- Select a list to CC --", "");
            ddl.Items.Insert(0, emptyItem);        
        }

        protected void chkGrpLotNo_CheckedChanged(object sender, EventArgs e)
        {
            if (chkGrpLotNo.Checked)
                rdlsNoOfLot.Enabled = true;
            else
                rdlsNoOfLot.Enabled = false;

        }

    

        

       
   }

   
}