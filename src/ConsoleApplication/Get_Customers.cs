using System;
using SAGE;
using ConsoleApplication;
using ExactOnline.Client.Sdk.Controllers;
using ExactOnline.Client.Models;

namespace SAGE
{
    public class Get_Customers
    {
        public Get_Customers(string companyName)
        {
            const string clientId = "b4b7919a-f51c-4272-9576-dbcccffb8764";
            const string clientSecret = "eWfc7W7BQXYZ";

            //This can be any url as long as it is identical to the callback url you specified for your app in the App Center.
            var callbackUrl = new Uri("http://websocketclient.local");

            var connector = new Connector(clientId, clientSecret, callbackUrl);
            var client = new ExactOnlineClient(connector.EndPoint, connector.GetAccessToken);

            //Declare Variables
            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();
            SageDataObject220.WorkSpace oWS;
            SageDataObject220.SalesRecord oSalesRecord;
            SageDataObject220.SalesDeliveryRecord oSalesDeliveryRecord;
            SageDataObject220.AgedBalances oAgedBalances;

            String szDataPath;

            //Instantiate WorkSpace
            oWS = (SageDataObject220.WorkSpace)oSDO.Workspaces.Add("Demodata");

            //Show select company dialog
            szDataPath = oSDO.SelectCompany(companyName);

            //Try a connection, will throw an exception if it fails
            try
            {
                //Leaving the username and password blank generates a login dialog
                oWS.Connect(szDataPath, "Manager", "", "Demodata");
                //Instantiate the Sales Record and Sales Delivery Record objects
                oSalesRecord = (SageDataObject220.SalesRecord)oWS.CreateObject("SalesRecord");
                oSalesDeliveryRecord = (SageDataObject220.SalesDeliveryRecord)oWS.CreateObject("SalesDeliveryRecord");
                var zxc = oSalesRecord.Count;
                // var zxcs = fffffff.Count;
                //Read the first supplier record
                oSalesRecord.MoveFirst();

                var cvb = oSalesDeliveryRecord.Count;
                oSalesDeliveryRecord.MoveFirst();

                oAgedBalances = (SageDataObject220.AgedBalances)oSalesRecord.GetAgedBalances();

                bool movNext2;
                string res = "";

                int cou = 0;

                do
                {
                    cou++;
                    //    new Guid("")
                    // Guid.TryParse();
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_MANAGER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_ON_HOLD") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_OPENED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_REF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_STATUS") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE_CUSTOMER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE_DONOR") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE_MEMBER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ADDRESS_1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ADDRESS_2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ADDRESS_3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ADDRESS_4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ADDRESS_5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ANALYSIS_1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ANALYSIS_2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "ANALYSIS_3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "COUNTRY_CODE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_APP_RECEIVED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_APPLIED_FOR") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_BF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_BUREAU") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_CARD_NO") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_CF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_LIMIT") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH10") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH11") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH12") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH6") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH7") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH8") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_MTH9") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_POSITION") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CREDIT_REFERENCE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "CURRENCY") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DATE_CREDIT_APP_RECEIVED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DECLARATION_VALID_FROM") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEF_NOM_CODE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEF_TAX_CODE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_ADDRESS_1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_ADDRESS_2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_ADDRESS_3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_ADDRESS_4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_ADDRESS_5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_CONTACT_NAME") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_FAX") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_NAME") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEL_TELEPHONE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DELETED_FLAG") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DEPT_NUMBER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DISCOUNT_RATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DISCOUNT_TYPE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DISPUTED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DONOR_FORENAME") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DONOR_SURNAME") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DONOR_TITLE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "DUNS_NUMBER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "E_MAIL") + ";";
                    res += SDOHelper.Read(oSalesRecord, "E_MAIL2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "E_MAIL3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "EC_CODE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "EXTERNAL_USAGE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "FAX") + ";";
                    res += SDOHelper.Read(oSalesRecord, "FIRST_HEADER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "FIRST_INVOICE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "GIFT_AID_ENABLED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "HOLD_MAIL") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INACTIVE_FLAG") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_BF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_CF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH6") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH7") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH8") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH9") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH10") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH11") + ";";
                    res += SDOHelper.Read(oSalesRecord, "INVOICE_MTH12") + ";";
                    res += SDOHelper.Read(oSalesRecord, "LAST_CREDIT_REV") + ";";
                    res += SDOHelper.Read(oSalesRecord, "LAST_HEADER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "LAST_INV_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "LAST_PAY_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "LAST_UPDATED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MEMBER_REMINDER_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MEMBER_RENEWAL_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MEMBER_SINCE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MEMO_OFFSET") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MEMO_SIZE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "MODIFY") + ";";
                    res += SDOHelper.Read(oSalesRecord, "NAME") + ";";
                    res += SDOHelper.Read(oSalesRecord, "NEXT_CREDIT_REV") + ";";
                    res += SDOHelper.Read(oSalesRecord, "NO_OF_HEADER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "OVERRIDE_PRODUCT_TAX") + ";";
                    res += SDOHelper.Read(oSalesRecord, "OVERRIDE_PRODUCT_NOMINAL") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_BF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_CF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_DUE_DAYS") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH6") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH7") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH8") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH9") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH10") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH11") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_MTH12") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PAYMENT_TYPE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PRICE_LIST_REF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PRIOR_YEAR") + ";";
                    res += SDOHelper.Read(oSalesRecord, "PRIORITY_TRADER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_BF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_CF") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH1") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH3") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH4") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH5") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH6") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH7") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH8") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH9") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH10") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH11") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECEIPT_MTH12") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECORD_CREATE_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECORD_DELETED") + ";";
                    res += SDOHelper.Read(oSalesRecord, "RECORD_MODIFY_DATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "SALES_REP") + ";";
                    res += SDOHelper.Read(oSalesRecord, "SEND_INVOICES_ELECTRONICALLY") + ";";
                    res += SDOHelper.Read(oSalesRecord, "SEND_LETTERS_ELECTRONICALLY") + ";";
                    res += SDOHelper.Read(oSalesRecord, "SETTLEMENT_DISC_RATE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "SETTLEMENT_DUE_DAYS") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TELEPHONE") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TELEPHONE_2") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TERMS") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TERMS_AGREED_FLAG") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TRADE_CONTACT") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TURNOVER_MTD") + ";";
                    res += SDOHelper.Read(oSalesRecord, "TURNOVER_YTD") + ";";
                    res += SDOHelper.Read(oSalesRecord, "VAT_REG_NUMBER") + ";";
                    res += SDOHelper.Read(oSalesRecord, "WWW") + ";";
                    res += Environment.NewLine;

                    Console.WriteLine(res);

                    if (cou > 2)
                    {
                        break;
                    }

                    Guid accountManager = new Guid();
                    Guid.TryParse(SDOHelper.Read(oSalesRecord, "ACCOUNT_MANAGER").ToString(), out accountManager);
                    bool blocked = new bool();
                    Boolean.TryParse(SDOHelper.Read(oSalesRecord, "ACCOUNT_ON_HOLD").ToString(), out blocked);
                    DateTime created = new DateTime();
                    DateTime.TryParse(SDOHelper.Read(oSalesRecord, "ACCOUNT_OPENED").ToString(),out created);
                    string remarks = SDOHelper.Read(oSalesRecord, "ACCOUNT_REF").ToString();
                    string type = SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE").ToString();
                    Guid businessType = new Guid();
                    Guid.TryParse(SDOHelper.Read(oSalesRecord, "ACCOUNT_TYPE_CUSTOMER").ToString(),out businessType);

                    string addressLine1 = SDOHelper.Read(oSalesRecord, "ADDRESS_1").ToString();
                    string addressLine2 = SDOHelper.Read(oSalesRecord, "ADDRESS_2").ToString();
                    string addressLine3 = SDOHelper.Read(oSalesRecord, "ADDRESS_3").ToString();
                    string city = SDOHelper.Read(oSalesRecord, "ADDRESS_4").ToString();
                    string country = SDOHelper.Read(oSalesRecord, "COUNTRY_CODE").ToString();

                    double creditLineSales = new double();
                    double.TryParse(SDOHelper.Read(oSalesRecord, "CREDIT_LIMIT").ToString(),out creditLineSales);

                    double discountSales = new double();
                    double.TryParse(SDOHelper.Read(oSalesRecord, "DISCOUNT_TYPE").ToString(),out discountSales);

                    string email = SDOHelper.Read(oSalesRecord, "E_MAIL").ToString();

                    if (email.IndexOf(';') != -1)
                    email = email.Remove(email.IndexOf(';'));

                    string fax = SDOHelper.Read(oSalesRecord, "FAX").ToString();

                    DateTime? endDate = null;
                    //DateTime.TryParse(SDOHelper.Read(oSalesRecord, "INACTIVE_FLAG").ToString(), out endDate);

                    if (SDOHelper.Read(oSalesRecord, "INACTIVE_FLAG").ToString() == "1")
                    {
                        endDate = new DateTime(2018, 01, 01);
                    }


                    string name = SDOHelper.Read(oSalesRecord, "NAME").ToString();

                    string paymentConditionSales = SDOHelper.Read(oSalesRecord, "PAYMENT_DUE_DAYS").ToString();

                    Account document = new Account
                    {
                        AccountManager = accountManager,
                        Blocked = blocked,
                        Created = created,
                        Remarks = remarks,
                        Type = type,
                        BusinessType = businessType,
                        AddressLine1 = addressLine1,
                        AddressLine2 = addressLine2,
                        AddressLine3 = addressLine3,
                        City = city,
                        // BankAccounts = 
                        // Contactname = 
                        Country = country,
                        CreditLineSales = creditLineSales,
                        DiscountSales = discountSales,
                        Email = email,
                        Fax = fax,
                        EndDate = endDate,
                        Name = name,
                        PaymentConditionSales = paymentConditionSales,
                        Status = "C",
                    };

                    bool createdClient = client.For<Account>().Insert(ref document);


                    movNext2 = oSalesRecord.MoveNext();

                } while (movNext2 == true);

                System.IO.File.WriteAllText(@"E:\Get_Customers.csv", res);

                oWS.Disconnect();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SDO Generated the Following Error: \n\n" + ex.Message, "Error!");
            }
        }
    }
}
