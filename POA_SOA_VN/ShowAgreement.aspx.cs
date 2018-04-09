using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
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
using CrystalDecisions.Reporting;
using CrystalDecisions.Shared;
using Microsoft.VisualBasic;
using ClosedXML.Excel;


namespace POA_SOA
{
    public partial class ShowAgreement : System.Web.UI.Page
    {
        static string OrderNumber;
        static string OrderType;
        string PA_Account;
        string BranchAdd;
        string PaymentTerm;
        float TaxRate;
        string Currency;
        bool USDOnly;
        bool isSaleOrder;
        float ExRate;
        string Branch;
        string oType;
        string DMSUser;
        string CustPO;
        int UseLotNo;
        static int SumQty;

        protected static SqlConnection oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
        protected static ReportDocument cryRpt = new ReportDocument();
        protected static DataTable POData;
        protected static SqlCommand cmd = new SqlCommand();
        
        
        protected void Page_Load(object sender, EventArgs e)
        {
            //ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Page load');", true);
            //cryRpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
           
        }
        protected void Page_Init(object sender, EventArgs e)
        {
            lblCCList.Text = "thuthuy@le-intl.com, erin@le-intl.com, huy@le-intl.com";
            lblToList.Text = "info.vn@le-intl.com";            
            
            CrystalReportViewer1.Visible = true;
            //btnSendEmail.Enabled = false;

            if (Request.Cookies["Info"] != null)
            {
                //Set variable from last page
                OrderNumber = (string)(Request.Cookies["Info"]["OrderNumber"]);
                OrderType = (string)(Request.Cookies["Info"]["OrderType"]);
                PA_Account = (string)(Request.Cookies["Info"]["Account"]);
                BranchAdd = (string)(Request.Cookies["Info"]["BranchAddress"]);
                PaymentTerm = (string)(Request.Cookies["Info"]["Term"]);
                Branch = (string)(Request.Cookies["Info"]["Branch"]);
                if (Request.Cookies["Info"]["USDOnly"] == "0") USDOnly = false;
                    else USDOnly = true;
                //if (Request.Cookies["Info"]["Username"].ToString().Trim()!= "") 
                //DMSUser = Request.QueryString["UserName"];
                DMSUser = (string)(Request.Cookies["Info"]["Username"]);
                Currency = (string)(Request.Cookies["Info"]["Curr"]);
                TaxRate = Convert.ToSingle((Request.Cookies["Info"]["TaxRate"]));//*100;
                CustPO = (string)(Request.Cookies["Info"]["CustPO"]);
                UseLotNo = Convert.ToByte(Request.Cookies["Info"]["UseLotNo"]);

                if ((Request.Cookies["Info"]["ExRate"] == null) || ((string)(Request.Cookies["Info"]["Taxable"]) == "N"))
                {
                    //TaxRate = 0;
                    ExRate = 0;
                }
                else
                {
                    //if (Branch=="B50") TaxRate = 10;
                    if ((!USDOnly) && (Branch == "B50")) ExRate = Convert.ToSingle(Request.Cookies["Info"]["ExRate"]);
                    else ExRate = 0;

                }

                PaymentTerm = (string)Request.Cookies["Info"]["Term"];

                if (Strings.Left(OrderType, 1).ToUpper() == "S") { isSaleOrder = true; }

                LoadReport();
                CrystalReportViewer1.ReportSource = cryRpt;
                CrystalReportViewer1.DataBind();

                //Control Sending email
                string sToList, sCCList;

                sToList = Request.Cookies["Info"]["ToList"].Trim();
                sCCList = Request.Cookies["Info"]["CCList"].Trim();

                if (sCCList == "")
                {
                    btnSendEmail.Enabled = false;
                    oSqlConnection.Close();
                    return;
                }
                else
                {
                    //btnSendEmail.Enabled = true;
                }
                //check validity of , in email lists
                if (sToList.Substring(0, 1) == ",") sToList = sToList.Substring(1, sToList.Length - 1);
                if (sToList.Substring(sToList.Length - 1, 1) == ",") sToList = sToList.Substring(0, sToList.Length - 1);
                if (sCCList.Substring(0, 1) == ",") sCCList = sCCList.Substring(1, sCCList.Length - 1);
                if (sCCList.Substring(sCCList.Length - 1, 1) == ",") sCCList = sCCList.Substring(0, sCCList.Length - 1);

                //add the Procurement list
                if ((OrderType == "OB") || (OrderType == "OP"))
                {
                    if (oSqlConnection.State == ConnectionState.Closed)
                        oSqlConnection.Open();
                    SqlDataAdapter dap = new SqlDataAdapter("SELECT TOList, CCList FROM [OPD_DB].[dbo].[SendingList] WHERE JDECode = 1", oSqlConnection);
                    DataTable dt = new DataTable();
                    dap.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        //Get data to variable 
                        sCCList += ", " + dt.Rows[0]["CCList"].ToString();
                    }
                }

                if (sCCList.Length > 2)
                {
                    if (sCCList.Substring(0, 1) == ",") sCCList = sCCList.Substring(1, sCCList.Length - 1);
                    if (sCCList.Substring(sCCList.Length - 1, 1) == ",") sCCList = sCCList.Substring(0, sCCList.Length - 1);
                }
                else sCCList = "";

                lblToList.Text = sToList;
                lblCCList.Text = sCCList;

                lblUsername.Text = DMSUser;

                //Get Subject 
                string SlCd1 = "", SlCd2 = "", ShipTo = "";
                try
                {
                    if (oSqlConnection.State == ConnectionState.Closed)
                        oSqlConnection.Open();
                    cmd = new SqlCommand("Proc_GetOrderInfo", oSqlConnection);
                    DataTable HeaderData = new DataTable();

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@OrderNo", SqlDbType.VarChar).Value = OrderNumber;
                    cmd.Parameters.Add("@OrderType", SqlDbType.VarChar).Value = OrderType;
                    cmd.ExecuteNonQuery();

                    SqlDataAdapter oSqlDataAdapter = new SqlDataAdapter(cmd);

                    oSqlDataAdapter.Fill(HeaderData);
                    SlCd1 = (string)(from DataRow dr in HeaderData.Rows
                                     select (string)dr["Brand"]).FirstOrDefault();

                    SlCd2 = (string)(from DataRow dr in HeaderData.Rows
                                     select (string)dr["ProductType"]).FirstOrDefault();

                    ShipTo = (string)(from DataRow dr in HeaderData.Rows
                                      select (string)dr["ShipToName"]).FirstOrDefault();
                }
                catch { }

                txtMailSubject.Text = SlCd1 + "-" + SlCd2 + "-" + ShipTo + ": " + CustPO + " - " + Branch + " - " + OrderNumber + " " + OrderType;

                oSqlConnection.Close();               
            
            }

            
            
            //testing trial
            /*string CCMailGroup;
            char[] sep = { ',', ';', ' ' };
            CCMailGroup = lblToList.Text.Split(sep).Last();
            lblCCList.Text = "amyhang@le-intl.com, thuthuy@le-intl.com, erin@le-intl.com, huy@le-intl.com";
            lblToList.Text = "info.vn@le-intl.com, " + CCMailGroup;
            txtMailSubject.Text += " - " +oType;
            txtEmailNote.Text = "TESTING ONLY";*/
        }


        protected void btnSendEmail_Click(object sender, EventArgs e)
        {
            string sNote, sMailSubject, sMailBody;
            lblUsername.Text = DMSUser;
            sNote = txtEmailNote.Text.Trim();
            sMailSubject = txtMailSubject.Text.Trim();
            System.Text.RegularExpressions.Regex.Replace(sMailSubject, @"<(.|\n)*?>", string.Empty);
            sMailBody = lblEmailBody.Text;

            if (sMailBody=="")
            {
                btnSendEmail.Enabled = false;
                lblEmailNotify.Text = "Error in mail content, please process again!";
            }

            lblEmailNotify.Text = emailReport("dms@le-intl.com", lblToList.Text.Trim(), lblCCList.Text.Trim(), sMailSubject, sNote, sMailBody);
            
            CrystalReportViewer1.Visible = false;
            lblEmailBody.Text = null;
            btnEmailBody.Enabled = false;
            btnSendEmail.Enabled = false;            
        }
        

        protected void btnGetPDF_Click(object sender, EventArgs e)
        {
            cryRpt.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Response, true, OrderType + OrderNumber);
        }
        

        protected void btnExportExcel_Click(object sender, EventArgs e)
        {
            exportExcel_Single(isSaleOrder);
        }

        protected void btnEmailBody_Click(object sender, EventArgs e)
        {
            lblEmailBody.Visible = true;            
            lblEmailBody.Text = MailBody("Send" + oType, txtEmailNote.Text, lblCCList.Text);
            btnSendEmail.Enabled = true;
            //lblTotalQty.Text = SumQty;
        }

        protected void SendEmail(string SenderAddress, string ToList, string CCList, string MSubject, string EmailNote, string MailContent)
        {
            //if (OrderNumber=="")

            //Generate PDF and Excel files
            cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, @"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".pdf");
            if (Strings.Left(OrderType, 1) == "O") exportExcel(false);



            if (ToList == "") return;

            string EmailAdd, Password;
            EmailAdd = "dms@le-intl.com";
            Password = "Qatu4394";

            string Subject;

            Subject = MSubject;
            //Body = MailBody("Send"+oType,EmailNote,CCList);
            //From email address as user

            using (MailMessage mm = new MailMessage(EmailAdd, ToList))
            {

                mm.CC.Add(CCList);
                mm.Subject = Subject;
                mm.Body = MailContent;//Body;
                //Add PDF file
                //mm.Attachments.Add(new Attachment(cryRpt.ExportToStream(ExportFormatType.PortableDocFormat), OrderType + OrderNumber+".pdf"));
                mm.Attachments.Add(new Attachment(@"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".pdf"));
                if (Strings.Left(OrderType, 1) == "O")
                    mm.Attachments.Add(new Attachment(@"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".xlsx"));
                AlternateView plainView = AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(mm.Body, @"<(.|\n)*?>", string.Empty), null, "text/plain");
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(mm.Body, null, "text/html");
                mm.AlternateViews.Add(plainView);
                mm.AlternateViews.Add(htmlView);
                mm.IsBodyHtml = true;
                mm.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");


                MailAddress mAdd = new MailAddress(EmailAdd);
                mm.Sender = mAdd;
                mAdd = new MailAddress(SenderAddress);
                mm.From = mAdd;


                //Send by dms@le-intl.com
                SmtpClient smtp = new SmtpClient("outlook.office365.com", 25);
                NetworkCredential NetworkCred = new NetworkCredential(EmailAdd, Password);
                smtp.Credentials = NetworkCred;
                //smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                smtp.Timeout = 180000;


                /*SmtpClient smtp = new SmtpClient();
                smtp.Host = "localhost";
                smtp.Port = 25;
                smtp.EnableSsl = false;*/


                //Log for email sending
                if (oSqlConnection.State == ConnectionState.Closed) oSqlConnection.Open();


                string tSQL;

                if (Strings.Left(OrderType, 1) == "O")
                    oType = "POA";
                else
                    oType = "SOA";

                //Sending email
                try
                {
                    smtp.Send(mm);

                    tSQL = "INSERT INTO dbo.SendingEmail_Log([OrderNumber],[OrderType],[JDEAddCode],[SendingTime],[Sender],[ToList],[CCList],[Note],[TypeofAttachment],[EmailSubject],[SumQty],[Currency]) VALUES('"
                       + OrderNumber + "','" + OrderType + "','"
                        + "'," //+ DateTime.Now
                        + "GETDATE(),'"
                        + DMSUser + "','"
                        + mm.To + "','"
                        + mm.CC + "','"
                        + EmailNote + "','"
                        + OrderType + OrderNumber + ".pdf"
                        //+ { foreach (Attachment att in mm.Attachments)  att.Name.ToString() + ";" }
                        + "','" + MSubject + "'," + SumQty + ",'" + Currency + "')";
                    //DMSUser = tSQL;
                    if (oSqlConnection.State == ConnectionState.Closed)
                        oSqlConnection.Open();
                    cmd = new SqlCommand(tSQL, oSqlConnection);
                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex) { lblEmailNotify.Text = ("Error sending:" + ex.Message); }
                finally
                {
                    if (oSqlConnection != null) oSqlConnection.Close();
                    if (cmd != null) oSqlConnection.Close();
                }

                lblEmailNotify.Text = "Mail sent successfully";

            }
        }
        void AddressInfo()
        {
            DataTable dtAccount = new DataTable();
            cmd = new SqlCommand("SELECT * FROM dbo.Address WHERE Code='" + BranchAdd + "'", oSqlConnection);
            SqlDataAdapter oSqlDataAdapter = new SqlDataAdapter(cmd);
            oSqlDataAdapter.Fill(dtAccount);

            string AddInfo = (string)(from DataRow dr in dtAccount.Rows
                              select (string)dr["Add1"]).FirstOrDefault();
            cryRpt.SetParameterValue("AddLine1", AddInfo);

            AddInfo = (string)(from DataRow dr in dtAccount.Rows
                              select (string)dr["Add2"]).FirstOrDefault();
            cryRpt.SetParameterValue("AddLine2", AddInfo);

            AddInfo = (string)(from DataRow dr in dtAccount.Rows
                              select (string)dr["Add3"]).FirstOrDefault();
            cryRpt.SetParameterValue("AddLine3", AddInfo);

            AddInfo = (string)(from DataRow dr in dtAccount.Rows
                      select (string)dr["Add4"]).FirstOrDefault();
            cryRpt.SetParameterValue("AddLine4", AddInfo);

            AddInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Note"]).FirstOrDefault();
            cryRpt.SetParameterValue("Note", AddInfo);

            //Set Company Name
            AddInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Name"]).FirstOrDefault();
            cryRpt.SetParameterValue("CompanyName", AddInfo);

        }
        

        void AccountInfo()
        {
            DataTable dtAccount = new DataTable();
            cmd = new SqlCommand("SELECT * FROM dbo.Account WHERE Account_Code='" + PA_Account + "'", oSqlConnection);
            SqlDataAdapter oSqlDataAdapter = new SqlDataAdapter(cmd);
            oSqlDataAdapter.Fill(dtAccount);

            string AccInfo = (from DataRow dr in dtAccount.Rows
                              select (string)dr["Account_Name"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_AccName", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Address"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_AccAddress", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Address_Line2"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_AccAdd2", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Bank_Name"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_BankName", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Bank_Address"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_BankAdd", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Account_USD"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_AccNo", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Account_LocalCur"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_AccNo2", AccInfo);

            AccInfo = (from DataRow dr in dtAccount.Rows
                       select (string)dr["Swift_Code"]).FirstOrDefault();
            cryRpt.SetParameterValue("AP_SwiftCode", AccInfo);
        
        }


       protected void GetPODataTable()
       {           
           if (Strings.Left(OrderType, 1) == "O")
               oType = "POA";
           else
               oType = "SOA";

           if (oSqlConnection.State == ConnectionState.Closed)
               oSqlConnection.Open();

           //Open command by Type of Report
           string ProcName = "Proc_SHOW_" + oType;
           if (((Branch == "B120") || (Branch == "B70")) && (Currency == "USD"))
               ProcName = "Proc_SHOW_" + oType + "_Foreign";
           if ((Branch == "B80") && (Currency == "IDR"))
               ProcName = "Proc_SHOW_"+ oType +"_Foreign";

           cmd = new SqlCommand(ProcName, oSqlConnection);
           POData = new DataTable();

           cmd.CommandType = CommandType.StoredProcedure;
           cmd.Parameters.Add("@OrderNo", SqlDbType.VarChar).Value = OrderNumber;
           cmd.Parameters.Add("@OrderType", SqlDbType.VarChar).Value = OrderType;
           cmd.ExecuteNonQuery();

           SqlDataAdapter oSqlDataAdapter = new SqlDataAdapter(cmd);

           oSqlDataAdapter.Fill(POData);
       }


        protected void LoadReport()
        {
            cryRpt.Close();
            cryRpt.Dispose();

            cryRpt = new ReportDocument();

            //Get Data Table
            GetPODataTable();
            

            //Get Report Path
            string RepPath = "~/Report/" ;
            if (oType=="POA")
                RepPath += oType;
            else 
                if (UseLotNo == 1)
                    if (USDOnly) RepPath = "~/Report/SOA_CustPO.rpt";
                    else RepPath = "~/Report/SOA_CustPO_VND.rpt";
                else if (UseLotNo == 2)
                    if (USDOnly) RepPath = "~/Report/SOA_CustPO2.rpt";
                    else RepPath = "~/Report/SOA_CustPO2_VND.rpt";

                else if (OrderType == "SX")
                    RepPath = "~/Report/SOA_SX";
                else
                    RepPath = "~/Report/SOA_SD";
           
            if (UseLotNo == 0)
            switch (Branch.ToUpper())
            {
                case "B40":
                    RepPath += ".rpt";                   
                    break;
                case "B50":
                    if (!USDOnly)
                        RepPath += "_VN.rpt";
                    else
                        RepPath += ".rpt"; 
                    break;
                case "B120":
                    if ((Currency == "INR") && (Strings.Right(OrderType,1) == "X"))
                        if (isSaleOrder)
                        {
                            RepPath = "~/Report/SOA_SX_INDIA.rpt";
                        }
                        else
                            RepPath = "~/Report/POA_OX_INDIA.rpt";
                    else if (Currency == "INR")
                        RepPath += "_INDIA.rpt";
                    else
                        RepPath += ".rpt";                    
                    break;
                case "B70":
                    if (isSaleOrder)
                    {
                        if (Currency == "THB")
                            RepPath = "~/Report/SOA_SD_THAI_THB.rpt";
                        else
                            RepPath += ".rpt";  
                    }
                    else
                        RepPath = "~/Report/POA_THAI.rpt";
                    break;
                case "B80":
                    if (isSaleOrder)
                    {
                        if ((OrderType == "SX") && (Currency == "IDR"))
                            RepPath = "~/Report/SOA_SX_INDO.rpt";
                        else if ((OrderType == "SD") && (Currency == "IDR"))
                            RepPath = "~/Report/SOA_SD_INDO.rpt";
                        else
                            RepPath += ".rpt";  
                    }
                    else
                        if (Currency == "IDR")
                            RepPath = "~/Report/POA_INDO.rpt";
                        else
                            RepPath += ".rpt";
                    break;
                case "B150":
                    if (isSaleOrder)
                    {
                        if ((OrderType == "SX") && (Currency == "IDR"))
                            RepPath = "~/Report/SOA_SX_INDO.rpt";
                        else if ((OrderType == "SD") && (Currency == "IDR"))
                            RepPath = "~/Report/SOA_SD_INDO.rpt";
                        else
                            RepPath += ".rpt";  
                    }
                    else
                        if (Currency == "IDR")
                            RepPath = "~/Report/POA_INDO.rpt";
                        else
                            RepPath += ".rpt";
                    break;

                default:
                    RepPath += ".rpt";          
                    break;
            }

            

            cryRpt.Load(Server.MapPath(RepPath));
            cryRpt.SetDataSource(POData);

            cryRpt.SetParameterValue("Term", PaymentTerm);

            if ((Branch == "B50") && (!USDOnly))
            {
                cryRpt.SetParameterValue("ExRate", ExRate);
            }
           
            cryRpt.SetParameterValue("TaxRate", TaxRate);
           
            if (isSaleOrder) AccountInfo();
            //Branch address for India
            AddressInfo();
            //if ((RepPath == "~/Report/SOA_SD.rpt") || (RepPath == "~/Report/POA_INDIA.rpt")) AddressInfo();     

            //Close connection
            oSqlConnection.Close();

        }

      string emailReport(string SenderAddress, string ToList, string CCList, string MSubject, string EmailNote, string MailContent)
        {
            //if (OrderNumber=="")
            
            //Generate PDF and Excel files
            cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, @"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".pdf");
            if (Strings.Left(OrderType, 1) == "O") exportExcel(false);



            if (ToList == "") return "No sending list!";

            string EmailAdd, Password;
            EmailAdd = "dms@le-intl.com";
            Password = "Qatu4394";

            string Subject;
            
            Subject = MSubject;
            //Body = MailBody("Send"+oType,EmailNote,CCList);
            //From email address as user

            using (MailMessage mm = new MailMessage(EmailAdd, ToList))
            {                

                mm.CC.Add(CCList);
                mm.Subject = Subject;
                mm.Body = MailContent;//Body;
                //Add PDF file
                //mm.Attachments.Add(new Attachment(cryRpt.ExportToStream(ExportFormatType.PortableDocFormat), OrderType + OrderNumber+".pdf"));
                mm.Attachments.Add(new Attachment(@"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".pdf"));
                if (Strings.Left(OrderType, 1) == "O")
                    mm.Attachments.Add(new Attachment(@"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".xlsx")); 
                AlternateView plainView = AlternateView.CreateAlternateViewFromString(System.Text.RegularExpressions.Regex.Replace(mm.Body, @"<(.|\n)*?>", string.Empty), null, "text/plain");
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(mm.Body, null, "text/html");
                mm.AlternateViews.Add(plainView);
                mm.AlternateViews.Add(htmlView);
                mm.IsBodyHtml = true;
                mm.BodyEncoding = System.Text.Encoding.GetEncoding("utf-8");


                MailAddress mAdd = new MailAddress(EmailAdd);
                mm.Sender = mAdd;
                mAdd = new MailAddress(SenderAddress);
                mm.From = mAdd;
                

                //Send by dms@le-intl.com
                SmtpClient smtp = new SmtpClient("outlook.office365.com", 25);
                NetworkCredential NetworkCred = new NetworkCredential(EmailAdd, Password);                
                smtp.Credentials = NetworkCred;
                //smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                smtp.Timeout = 180000;


                /*SmtpClient smtp = new SmtpClient();
                smtp.Host = "localhost";
                smtp.Port = 25;
                smtp.EnableSsl = false;*/


                //Log for email sending
                if (oSqlConnection.State == ConnectionState.Closed) oSqlConnection.Open();

                
                string tSQL;

                if (Strings.Left(OrderType, 1) == "O")
                    oType = "POA";
                else
                    oType = "SOA";

                //Sending email
                try 
                {
                    smtp.Send(mm);

                    tSQL = "INSERT INTO dbo.SendingEmail_Log([OrderNumber],[OrderType],[JDEAddCode],[SendingTime],[Sender],[ToList],[CCList],[Note],[TypeofAttachment],[EmailSubject],[SumQty],[Currency]) VALUES('"
                       + OrderNumber + "','" + OrderType + "','"
                        + "'," //+ DateTime.Now
                        + "GETDATE(),'"
                        + DMSUser + "','"
                        + mm.To + "','"
                        + mm.CC + "','"
                        + EmailNote + "','"
                        + OrderType + OrderNumber + ".pdf"
                        //+ { foreach (Attachment att in mm.Attachments)  att.Name.ToString() + ";" }
                        + "','" + MSubject + "'," + SumQty +",'" + Currency + "')";
                    //DMSUser = tSQL;
                    if (oSqlConnection.State == ConnectionState.Closed)
                        oSqlConnection.Open();
                    cmd = new SqlCommand(tSQL, oSqlConnection);
                    cmd.ExecuteNonQuery();
                
                }
                catch (Exception ex) { return ("Error sending:" + ex.Message); }
                finally
                {
                    if (oSqlConnection != null) oSqlConnection.Close();
                    if (cmd != null) oSqlConnection.Close();
                }

                return "Mail sent successfully";
            }
            
        }




      string MailBody(string MailType, string MailNote, string CCMail)
        {
            string mBody, tSQL;
            SqlDataAdapter oSqlDataAdapter;
            DataTable table ;
            

            //CCMail cut out whole string and get mailgroup only
            string CCMailGroup;
            char[] sep = { ',',';',' ' };
            CCMailGroup = CCMail.Split(sep)[0];


            
            switch (MailType)
            {
                case "SendPOA":
                    mBody = "<meta http-equiv='Content-Type' content='text/html; charset=us-ascii'>  <style type='text/css'> .style10 {color: #FFFFFF;background-color: #0000FF;}</style>"
                        + "<span style='font-size: x-large;font-weight: bold;color: #FF0000;'>***POA***</span><br><br>"
                        + "Dear Partnership Supplier,<br><br> Please find attachment of our order " + OrderType + " " + OrderNumber + " to you. Please kindly confirm back your <i><u>verification</u></i> and <i><u>receipt</u></i> to us once safely receiving. <br>"
                        + "<br> Our Planning team (DR Team) will send you the delivery request via our Delivery Management System (DMS), therefore your high attention and <b><u>follow-up on DMS</b></u> is highly appreciated."
                        + "<br><br><b><span style='background-color: #FFFF00'>NOTE: " + MailNote + "</span>"
                        + "<br><br>The Order Details as below: </b><br><br>"
                        + "<table cellspacing='0' cellpadding='4' rules='all' border='1' style='border-color:#3366CC;border-width:1px;border-collapse:collapse;'>";
                    mBody += " <tr> <td class='style10'>NO.</td>"
                        + "<td class='style10'>ITEM CODE</td>"
                        + "<td class='style10'>DESCRIPTION</td>"
                        + "<td class='style10'>QTY.</td>"
                        //+ "<td class='style10'>UNIT COST</td>"
                        //+ "<td class='style10'>AMOUNT</td>"
                        + "<td class='style10'>REQUEST DATE</td>"
                        //+ "<td class='style10'>BLANKET ORDER</td>"
                        + "<td class='style10'>CUSTOMER PO</td>"
                        + "<td class='style10'>S.O. NO</td>"
                        + "<td class='style10'>S.O. TYPE</td></tr>";


                        foreach (DataRow row in POData.Select())
			            {
                            mBody+= "<tr>";
                            mBody += "<td>" + (int)Convert.ToDecimal(row["Line_No"].ToString()) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["ItemCode"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["Description"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["QtyOrdered"]) + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["UnitCost"]) + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["Amount"]) + "</td>";
                            mBody += "<td>" + DateTime.Parse(row["RequestDate"].ToString()).ToString("MMM-dd-yy") + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["Blanket"]) + "</td>";
                            if (oType=="POA")
                                mBody += "<td>" + Convert.ToString(row["CustomerPO"]) + "</td>";
                            else
                                mBody += "<td>" + Convert.ToString(row["Reference1"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["SalesOrderNo"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["SalesOrderType"]) + "</td>";
                            mBody += "</tr>";

                            //Write order detail log here
                            string CustPO;
                            float UnitCost,UnitPrice;
                            if (oType == "POA")
                            {
                                CustPO = row["CustomerPO"].ToString();
                                UnitCost = Convert.ToSingle(row["UnitCost"]);
                                UnitPrice = 0;
                            }
                            else
                            {
                                CustPO = row["Reference1"].ToString();
                                UnitCost = 0;
                                UnitPrice = Convert.ToSingle(row["UnitPrice"]);
                            }


                            tSQL = "INSERT INTO PrintingOrder_Detail_Log([EmailSentTo_JDECode],[SendingTime],[Sender],[OrderNumber],[OrderType],[Line_No],[ItemCode],[Description1],[CustomerPO],[Blanket] "
                                    + ",[Taxable],[TaxCode1],[TaxCode2],[QtyOrdered],[UnitCost],[UnitPrice],[RequestDate],[CurrencyCode],[JDEUserID]) VALUES('"
                                    + row["SupplierCode"].ToString() + "',"
                                    + "GETDATE(),'" + DMSUser + "','"
                                    + OrderNumber + "','" + OrderType + "','"
                                    + row["Line_No"].ToString() + "','"
                                    + row["ItemCode"].ToString() + "','"
                                    + row["Description"].ToString() + "','"
                                    + CustPO + "','"
                                    + row["Blanket"].ToString() + "','"
                                    + row["Taxable"].ToString() + "','"
                                    + row["TaxCode1"].ToString() + "','"
                                    + row["TaxCode2"].ToString() + "',"
                                    + row["QtyOrdered"] + ","
                                    + UnitCost + "," + UnitPrice + " ,'"
                                    + row["RequestDate"] + "','" + Currency +"','"
                                    + row["JDEUserID"].ToString() + "')";

                            if (oSqlConnection.State == ConnectionState.Closed) oSqlConnection.Open();
                            cmd = new SqlCommand(tSQL, oSqlConnection);
                            cmd.ExecuteNonQuery();
                            
                        }

                        //Input SUM qty
                        tSQL = "EXEC [dbo].[Proc_SumQty] '" + OrderNumber + "', '" + OrderType + "'";

                        cmd = new SqlCommand(tSQL, oSqlConnection);
                        cmd.ExecuteNonQuery();
                        oSqlDataAdapter = new SqlDataAdapter(cmd);
                        table = new DataTable();
                        oSqlDataAdapter.Fill(table);
                        SumQty = (int)table.Rows[0]["SumQty"];

                        //string SumQ = Convert.ToString(table.Rows[0]["SumQty"]);
                        mBody += "<tr><td></td><td></td><td><b> SUM: </b></td>";
                        mBody += "<td><b>" + Convert.ToString(SumQty) + "</b></td><td></td><td></td><td></td><td></td></tr>";
                        
                        mBody += "";
                        oSqlConnection.Close();


                        mBody += "</table><br><br><br>Should you have question or further clarification on the order, please do not hesitate to contact me.<br>";

                    break;

                case "SendSOA":   
            
                    mBody = "<meta http-equiv='Content-Type' content='text/html; charset=us-ascii'>  <style type='text/css'> .style10 {color: #FFFFFF;background-color: #0000FF;}</style>"
                        + "Dear Valued Customers,<br><br>Thank you so much for your order to us.<br><br> Please kindly find the Sales Acknowledgement Agreement (SOA) file of the summary of your order with the confirmation number " + OrderType + " " + OrderNumber + " from us. <br>"
                        + "<br><br><span style='background-color: #FFFF00'><b>NOTE: " + MailNote + "</span>"
                        + "<br><br>Details of your order as below: </b><br><br>"
                        + "<table cellspacing='0' cellpadding='4' rules='all' border='1' style='border-color:#3366CC;border-width:1px;border-collapse:collapse;'>";
                    mBody += " <tr> <td class='style10'>NO.</td>"
                        + "<td class='style10'>ITEM CODE</td>"
                        + "<td class='style10'>DESCRIPTION</td>"
                        + "<td class='style10'>QTY.</td>"
                        //+ "<td class='style10'>UNIT COST</td>"
                        //+ "<td class='style10'>AMOUNT</td>"
                        + "<td class='style10'>REQUEST DATE</td>"
                        //+ "<td class='style10'>BLANKET ORDER</td>"
                        + "<td class='style10'>CUSTOMER PO</td>";
                        foreach (DataRow row in POData.Select())
			            {
                            mBody+= "<tr>";
                            mBody += "<td>" + (int)Convert.ToDecimal(row["Line_No"].ToString()) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["ItemCode"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["Description"]) + "</td>";
                            mBody += "<td>" + Convert.ToString(row["QtyOrdered"]) + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["UnitCost"]) + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["Amount"]) + "</td>";
                            mBody += "<td>" + DateTime.Parse(row["RequestDate"].ToString()).ToString("MMM-dd-yy") + "</td>";
                            //mBody += "<td>" + Convert.ToString(row["Blanket"]) + "</td>";
                            if (oType=="POA")
                                mBody += "<td>" + Convert.ToString(row["CustomerPO"]) + "</td>";
                            else
                                mBody += "<td>" + Convert.ToString(row["Reference1"]) + "</td>";
                            mBody+= "</tr>";

                            //Write order detail log here
                            string CustPO;
                            float UnitCost,UnitPrice;
                            if (oType == "POA")
                            {
                                CustPO = row["CustomerPO"].ToString();
                                UnitCost = Convert.ToSingle(row["UnitCost"]);
                                UnitPrice = 0;
                            }
                            else
                            {
                                CustPO = row["Reference1"].ToString();
                                UnitCost = 0;
                                UnitPrice = Convert.ToSingle(row["UnitPrice"]);
                            }

                            tSQL = "INSERT INTO PrintingOrder_Detail_Log([EmailSentTo_JDECode],[SendingTime],[Sender],[OrderNumber],[OrderType],[Line_No],[ItemCode],[Description1],[CustomerPO],[Blanket] "
                                    + ",[Taxable],[TaxCode1],[TaxCode2],[QtyOrdered],[UnitCost],[UnitPrice],[RequestDate],[CurrencyCode],[JDEUserID]) VALUES('"
                                    + row["SupplierCode"].ToString() + "',"
                                    + "GETDATE(),'" + DMSUser + "','"
                                    + OrderNumber + "','" + OrderType + "','"
                                    + row["Line_No"].ToString() + "','"
                                    + row["ItemCode"].ToString() + "','"
                                    + row["Description"].ToString() + "','"
                                    + CustPO + "','"
                                    + row["Blanket"].ToString() + "','"
                                    + row["Taxable"].ToString() + "','"
                                    + row["TaxCode1"].ToString() + "','"
                                    + row["TaxCode2"].ToString() + "',"
                                    + row["QtyOrdered"] + ","
                                    + UnitCost + "," + UnitPrice + " ,'"
                                    + row["RequestDate"] + "','" + Currency +"','"
                                    + row["JDEUserID"].ToString() + "')";

                            if (oSqlConnection.State == ConnectionState.Closed) oSqlConnection.Open();
                            cmd = new SqlCommand(tSQL, oSqlConnection);
                            cmd.ExecuteNonQuery();
                            
                        }

                        //Input SUM qty
                        tSQL = "EXEC [dbo].[Proc_SumQty] '" + OrderNumber + "', '" + OrderType + "'";

                        cmd = new SqlCommand(tSQL, oSqlConnection);
                        cmd.ExecuteNonQuery();
                        oSqlDataAdapter = new SqlDataAdapter(cmd);
                        table = new DataTable();
                        oSqlDataAdapter.Fill(table);
                        SumQty = (int)table.Rows[0]["SumQty"];

                        mBody += "<tr><b><td></td><td></td><td><b> SUM: </b></td>";
                        mBody += "<td><b>" + Convert.ToString(SumQty) + "</b></td><td></td><td></td></tr>";

                        mBody += "";
                        oSqlConnection.Close();

                        mBody += "</table><br><br><b>Highlight side note:</b> we will start producing your order once we receive your confirmation on our Sales Acknowledgement Agreement.<br><br>Should you have any question or further request, please do not hesitate to contact us.<br>";

                    break;
                default:
                   
                    mBody="";
                    break;
            }
            mBody += "<br><br>L&amp;E Customer Service Department<br><br>" + DMSUser;
            mBody += "<span style='background-color: #FFFF00'><br><br>*This mail is automatically sent out from our system, please DO NOT reply to this email address.<br /> If you have any questions, please contact our team using " + CCMailGroup;
            mBody += "<br><br>*Email này được gửi từ hệ thống tự động của chúng tôi, xin vui lòng không trả lời bằng địa chỉ mail này. <br />Nếu bạn có bất kỳ thắc mắc nào xin vui lòng liên hệ chúng tôi theo địa chỉ mail :" + CCMailGroup;
            mBody += "<br><br>*此郵件從系統自動發出, 請勿用此郵件回復. 若有疑問煩請使用 " + CCMailGroup + " 與我司聯繫. 謝謝.</span>";  

            return mBody;
        }

        protected void exportExcel_Single(bool withPrice)
        {
            MemoryStream MyMemoryStream = new MemoryStream();
            string tSQL, oType;

            if (Strings.Left(OrderType, 1) == "O")
                oType = "POA";
            else
                oType = "SOA";

            tSQL = "EXEC [dbo].[Proc_SHOW_" + oType + "_Excel] '" + OrderNumber + "', '" + OrderType + "'";
            SqlConnection oSqlConnection = null;
            SqlCommand cmd = null;
            try
            {
                //Get DataTable
                oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
                oSqlConnection.Open();
                cmd = new SqlCommand(tSQL, oSqlConnection);

                SqlDataAdapter oAdapter = new SqlDataAdapter();
                oAdapter.SelectCommand = cmd;

                DataTable table = new DataTable(oType + "-" + OrderType + OrderNumber);
                oAdapter.Fill(table);

                //Remove columns of POA
                string[] removeCols = { };

                if (oType == "POA")
                {
                    removeCols = new string[] { "ExchangeRate", "ForeignRoundedUnitPrice", "USDUnitPrice", "USDPriceAmount" };

                }
                else if (!withPrice)
                {
                    removeCols = new string[] { "UnitPrice", "Amount", "ForeignUnitPrice", "UnitCost", "ExchangeRate", "ForeignRoundedUnitPrice", "USDUnitPrice", "USDPriceAmount" };
                }

                try
                {
                    foreach (string col in removeCols)
                        table.Columns.Remove(col);
                } catch { }


                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(table);

                    Response.Clear();
                    Response.Buffer = true;
                    Response.Charset = "";
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename=" + OrderType + OrderNumber + ".xlsx");
                    
                    
                    using (MyMemoryStream = new MemoryStream())
                    {
                        wb.SaveAs(MyMemoryStream);                        
                        MyMemoryStream.WriteTo(Response.OutputStream);
                        Response.Flush();
                        Response.End();
                        Response.Clear();
                    }

                }



            }
            catch (Exception ex)
            {
                //ex.Message;
            }
            finally
            {
                if (oSqlConnection != null) oSqlConnection.Dispose();
                if (cmd != null) oSqlConnection.Dispose();
            }

        }

        protected void exportExcel_Single1()
        {
            string filePath;
            filePath = @"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".xlsx";

            if (File.Exists(Server.MapPath(filePath)))
            {
                Response.Clear();
                Response.Buffer = true;
                Response.ContentType = "application/octet-stream";
                Response.AddHeader("content-disposition", "attachment;filename=" + OrderType + OrderNumber + ".xlsx");
                Response.TransmitFile(Server.MapPath(filePath));

                Response.Flush();
                Response.End();
            }
                

        }


        protected void exportExcel(bool withPrice)
        {
           string tSQL, oType;

            if (Strings.Left(OrderType, 1) == "O")
                oType = "POA";
            else
                oType = "SOA";

            tSQL = "EXEC [dbo].[Proc_SHOW_" + oType + "_Excel] '" + OrderNumber + "', '" + OrderType + "'";
            SqlConnection oSqlConnection = null;
            SqlCommand cmd = null;
            try
            {
                //Get DataTable
                oSqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["OPD_DBConnectionString1"].ConnectionString);
                oSqlConnection.Open();
                cmd = new SqlCommand(tSQL, oSqlConnection);

                SqlDataAdapter oAdapter = new SqlDataAdapter();
                oAdapter.SelectCommand = cmd;

                DataTable table = new DataTable(oType + "-" + OrderType + OrderNumber);
                oAdapter.Fill(table);

                //Remove columns of POA
                string[] removeCols = {};

                if (oType == "POA")
                {
                    removeCols = new string[] { "ExchangeRate", "ForeignRoundedUnitPrice", "USDUnitPrice", "USDPriceAmount" };
                
                }
                else if (!withPrice)     
                {
                    removeCols = new string[] { "UnitPrice", "Amount", "ForeignUnitPrice", "UnitCost", "ExchangeRate", "ForeignRoundedUnitPrice", "USDUnitPrice", "USDPriceAmount" };
                }                    

                
                try
                {
                    foreach (string col in removeCols)
                        table.Columns.Remove(col);                    
                }
                catch (Exception ex){}


                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(table);
                    wb.SaveAs(@"c:\temp\vietnam\poa-soa\" + OrderType + OrderNumber + ".xlsx");                    
                }               

            }
            catch (Exception ex)
            {
                //ex.Message;
            }
            finally
            {
                if (oSqlConnection != null) oSqlConnection.Dispose();
                if (cmd != null) oSqlConnection.Dispose();
            }
            
        }

    }

}