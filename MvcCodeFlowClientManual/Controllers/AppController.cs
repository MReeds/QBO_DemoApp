using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Net;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using System.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Windows.Forms;

namespace MvcCodeFlowClientManual.Controllers
{
    public class AppController : Controller
    {
        public static string clientid = ConfigurationManager.AppSettings["clientid"];
        public static string clientsecret = ConfigurationManager.AppSettings["clientsecret"];
        public static string redirectUrl = ConfigurationManager.AppSettings["redirectUrl"];
        public static string environment = ConfigurationManager.AppSettings["appEnvironment"];

        public static OAuth2Client auth2Client = new OAuth2Client(clientid, clientsecret, redirectUrl, environment);

        /// Use the Index page of App controller to get all endpoints from discovery url
        public ActionResult Index()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            Session.Clear();
            Session.Abandon();
            Request.GetOwinContext().Authentication.SignOut("Cookies");
            return View();
        }

        /// Start Auth flow
        public ActionResult InitiateAuth(string submitButton)
        {
            switch (submitButton)
            {
                case "Connect to QuickBooks":
                    List<OidcScopes> scopes = new List<OidcScopes>();
                    scopes.Add(OidcScopes.Accounting);
                    string authorizeUrl = auth2Client.GetAuthorizationURL(scopes);
                    return Redirect(authorizeUrl);
                default:
                    return (View());
            }
        }

        public ServiceContext CreateServiceContext()
        {
            string realmId = Session["realmId"].ToString();

            var principal = User as ClaimsPrincipal;
            OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);

            // Create a ServiceContext with Auth tokens and realmId
            ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
            serviceContext.IppConfiguration.MinorVersion.Qbo = "23";

            return serviceContext;
        }
        

        /// QBO API Request
        public ActionResult ApiCallService()
        {
            if (Session["realmId"] != null)
            {
                try
                {
                    var serviceContext = CreateServiceContext();

                    // Create a QuickBooks QueryService using ServiceContext
                    QueryService<CompanyInfo> querySvc = new QueryService<CompanyInfo>(serviceContext);
                    CompanyInfo companyInfo = querySvc.ExecuteIdsQuery("SELECT * FROM CompanyInfo").FirstOrDefault();

                    string output = "Company Name: " 
                        + companyInfo.CompanyName 
                        + " Company Address: " 
                        + companyInfo.CompanyAddr.Line1 
                        + ", " + companyInfo.CompanyAddr.City 
                        + ", " + companyInfo.CompanyAddr.Country 
                        + " " + companyInfo.CompanyAddr.PostalCode;

                    return View("ApiCallService", (object) output);
                }
                catch (Exception ex)
                {
                    return View("ApiCallService", (object)("QBO API call Failed!" + " Error message: " + ex.Message));
                }
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }

        public string BuildCustomerString(Customer customer)
        {
            StringBuilder sbCustomer = new StringBuilder();
            sbCustomer.Append("Customer Name: " + customer.GivenName + " " + customer.FamilyName);
            sbCustomer.AppendLine("\n");

            sbCustomer.AppendLine(customer.PrimaryPhone.FreeFormNumber);
            sbCustomer.AppendLine("\n");

            sbCustomer.Append("Display Name: " + customer.FullyQualifiedName);
            sbCustomer.AppendLine("\n");

            sbCustomer.Append("Billing Address: " + customer.BillAddr.Line1
                + customer.BillAddr.City + ","
                + " " + customer.BillAddr.CountrySubDivisionCode
                + " " + customer.BillAddr.PostalCode);

            string output = sbCustomer.ToString();

            return output; 
        }

        public ActionResult GetCustomerInfo(string customerName, bool checkDownload = false)
        {
            if (Session["realmId"] != null)
            {
                try
                {
                    var serviceContext = CreateServiceContext();

                    // Create a QuickBooks QueryService using ServiceContext
                    var selectStatement = $"Select * FROM Customer c WHERE c.DisplayName LIKE '%{customerName}%'";

                    QueryService<Customer> queryCustomer = new QueryService<Customer>(serviceContext);
                    Customer customer = queryCustomer.ExecuteIdsQuery(selectStatement).FirstOrDefault();

                    string output = BuildCustomerString(customer);

                    if(checkDownload == true)
                    {
                        var jsonOutput = JsonConvert.SerializeObject(customer);
                        CreateExcelWorkbook(customer, jsonOutput);
                    }

                    return View("ApiCustomer", (object) output);
                }
                catch (Exception ex)
                {
                    return View("ApiCustomer", (object)("QBO API call Failed!" + " Error message: " + ex.Message));
                }
            }
            else
                return View("ApiCustomer", (object)"QBO API call Failed!");
        }

        public void CreateExcelWorkbook(Customer customerObject, string customer)
        {
            // Create a Workbook object
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set JsonLayoutOptions
            JsonLayoutOptions options = new JsonLayoutOptions();
            options.ArrayAsTable = true;

            // Import JSON Data
            JsonUtility.ImportData(customer, worksheet.Cells, 0, 0, options);

            var customerName = customerObject.GivenName + "-" + customerObject.FamilyName;

            // Save Excel file
            workbook.Save($@"H:\repos\CSharp\{customerName}-QB-File.xlsx");
        }

        /// Use the Index page of App controller to get all endpoints from discovery url
        public ActionResult Error()
        {
            return View("Error");
        }

        /// Action that takes redirection from Callback URL
        public ActionResult Tokens()
        {
            return View("Tokens");
        }
    }
}