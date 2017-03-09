using ConsoleApplication;
using ExactOnline.Client.Models;
using ExactOnline.Client.Sdk.Controllers;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SAGE
{
    public class Get_Suppliers
    {
        public Get_Suppliers(string companyName)
        {

            //const string clientId = "b4b7919a-f51c-4272-9576-dbcccffb8764";
            //const string clientSecret = "eWfc7W7BQXYZ";

            ////This can be any url as long as it is identical to the callback url you specified for your app in the App Center.
            //var callbackUrl = new Uri("http://websocketclient.local");

            //var connector = new Connector(clientId, clientSecret, callbackUrl);
            //var client = new ExactOnlineClient(connector.EndPoint, connector.GetAccessToken);

            //Declare Objects
            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();
            SageDataObject220.WorkSpace oWS;
            SageDataObject220.SupplierAddress oSupplierAddress;
            SageDataObject220.PurchaseRecord oPurchaseRecord;

            //Decalre Variables
            String szDataPath;

            //Create the SDO Engine Object
            oSDO = new SageDataObject220.SDOEngine();

            //Create the Workspace
            oWS = (SageDataObject220.WorkSpace)oSDO.Workspaces.Add("Example");

            //Select company. The select company method taked the program install folder
            //as a parameter
            szDataPath = oSDO.SelectCompany(companyName);
            string res = "";

            if (szDataPath != string.Empty)
            {
                //Connect to the Data Files
                if (oWS.Connect(szDataPath, "Manager", "", "Demodata"))
                {
                    //Create an instance of the PurchaseRecord
                    oPurchaseRecord = (SageDataObject220.PurchaseRecord)oWS.CreateObject("PurchaseRecord");
                    var oPurchaseRecordCount = oPurchaseRecord.Count;
                    //Move to the first PurchaseRecord
                    oPurchaseRecord.MoveFirst();
                    int count = 0;
                    //Loop through the Sales Records

                    do
                    {
                        count++;

                        if (SDOHelper.Read(oPurchaseRecord, "INACTIVE_FLAG").ToString() == "1")
                        {
                            continue;
                        }

                        //res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_MANAGER") + ";";
                        res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_ON_HOLD") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_OPENED") + ";";
                        res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_REF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_STATUS") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ACCOUNT_TYPE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ANALYSIS_1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ANALYSIS_2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "ANALYSIS_3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BACS") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BACS_REFERENCE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BALANCE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ACCOUNT_NAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ACCOUNT_NUMBER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDITIONAL_REF1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDITIONAL_REF2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDITIONAL_REF3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDRESS_1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDRESS_2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDRESS_3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDRESS_4") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ADDRESS_5") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_BICSWIFT") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_IBAN") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_NAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_ROLL_NUMBER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BANK_SORT_CODE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BUILD_SOCIETY_REF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "BUILDING_SOCIETY") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CAN_APPLY_CHARGES") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CONTACT_NAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "COUNTRY_CODE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_APP_RECEIVED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_APPLIED_FOR") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_BF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_BUREAU") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_CARD_NO") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_CF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_LIMIT") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH10") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH11") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH12") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH4") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH5") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH6") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH7") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH8") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_MTH9") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_POSITION") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CREDIT_REFERENCE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "CURRENCY") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DATE_CREDIT_APP_RECEIVED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DECLARATION_VALID_FROM") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DEF_NOM_CODE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DEF_TAX_CODE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DELETED_FLAG") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DEPT_NUMBER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DISCOUNT_RATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DISCOUNT_TYPE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DISPUTED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DONOR_FORENAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DONOR_SURNAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DONOR_TITLE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "DUNS_NUMBER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "E_MAIL") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "E_MAIL2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "E_MAIL3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "EC_CODE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "EXTERNAL_USAGE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "FAX") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "FIRST_HEADER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "FIRST_INVOICE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "GIFT_AID_ENABLED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "HOLD_MAIL") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INACTIVE_FLAG") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_BF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_CF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH10") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH11") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH12") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH4") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH5") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH6") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH7") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH8") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "INVOICE_MTH9") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "LAST_CREDIT_REV") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "LAST_HEADER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "LAST_INV_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "LAST_PAY_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "LAST_UPDATED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "MEMBER_REMINDER_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "MEMBER_RENEWAL_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "MEMBER_SINCE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "MEMO_OFFSET") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "MEMO_SIZE") + ";";
                        res += SDOHelper.Read(oPurchaseRecord, "MODIFY") + ";";
                        res += SDOHelper.Read(oPurchaseRecord, "NAME") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "NEXT_CREDIT_REV") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "NO_OF_HEADER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "OVERRIDE_PRODUCT_NOMINAL") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "OVERRIDE_PRODUCT_TAX") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_BF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_CF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_DUE_DAYS") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH10") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH11") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH12") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH4") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH5") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH6") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH7") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH8") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_MTH9") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PAYMENT_TYPE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PRICE_LIST_REF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PRIOR_YEAR") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "PRIORITY_TRADER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_BF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_CF") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH1") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH3") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH4") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH5") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH6") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH7") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH8") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH9") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH10") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH11") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECEIPT_MTH12") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECORD_CREATE_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECORD_DELETED") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "RECORD_MODIFY_DATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "SALES_REP") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "SEND_INVOICES_ELECTRONICALLY") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "SEND_LETTERS_ELECTRONICALLY") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "SETTLEMENT_DISC_RATE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "SETTLEMENT_DUE_DAYS") + ";";
                        res += SDOHelper.Read(oPurchaseRecord, "TELEPHONE") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TELEPHONE_2") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TERMS") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TERMS_AGREED_FLAG") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TRADE_CONTACT") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TURNOVER_MTD") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "TURNOVER_YTD") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "VAT_REG_NUMBER") + ";";
                        //res += SDOHelper.Read(oPurchaseRecord, "WWW") + ";";
                        //res += Environment.NewLine;

                        Console.WriteLine(res);

                        if (count > 2)
                        {
                            break;
                        }

                    //    string accountManagerFullName = SDOHelper.Read(oPurchaseRecord, "ACCOUNT_MANAGER").ToString();
                    //    bool blocked = new bool();
                    //    bool.TryParse(SDOHelper.Read(oPurchaseRecord, "ACCOUNT_ON_HOLD").ToString(), out blocked);
                    //    DateTime created = new DateTime();
                    //    DateTime.TryParse(SDOHelper.Read(oPurchaseRecord, "ACCOUNT_OPENED").ToString(), out created);
                    //    string type = SDOHelper.Read(oPurchaseRecord, "ACCOUNT_TYPE").ToString();

                    //    List<BankAccount> bankAccount = new List<BankAccount>(); //SDOHelper.Read(oPurchaseRecord, "BANK_ACCOUNT_NAME")
                    //    double creditLinePurchase = new double();
                    //    double.TryParse(SDOHelper.Read(oPurchaseRecord, "CREDIT_LIMIT").ToString(), out creditLinePurchase);

                    //    string currency = SDOHelper.Read(oPurchaseRecord, "CURRENCY").ToString();

                    //    Guid gLAccountPurchase = new Guid();
                    //    Guid.TryParse(SDOHelper.Read(oPurchaseRecord, "DEF_NOM_CODE").ToString(), out gLAccountPurchase);

                    //    double discountPurchase = new double();
                    //    double.TryParse(SDOHelper.Read(oPurchaseRecord, "DISCOUNT_RATE").ToString(), out discountPurchase);

                    //    string email = SDOHelper.Read(oPurchaseRecord, "E_MAIL").ToString();
                    //    if (email.IndexOf(';') != -1)
                    //        email = email.Remove(email.IndexOf(';'));

                    //    string fax = SDOHelper.Read(oPurchaseRecord, "FAX").ToString();


                    //    DateTime? endDate = null;
                    //    //DateTime.TryParse(SDOHelper.Read(oSalesRecord, "INACTIVE_FLAG").ToString(), out endDate);
                    //    if (SDOHelper.Read(oPurchaseRecord, "INACTIVE_FLAG").ToString() == "1")
                    //    {
                    //        endDate = new DateTime(2018, 01, 01);
                    //    }

                    //    string name = SDOHelper.Read(oPurchaseRecord, "NAME").ToString();

                    //    string paymentConditionPurchase = SDOHelper.Read(oPurchaseRecord, "PAYMENT_DUE_DAYS").ToString();

                    //    int invoicingMethod = new int();
                    //    int.TryParse(SDOHelper.Read(oPurchaseRecord, "SEND_INVOICES_ELECTRONICALLY").ToString(), out invoicingMethod);

                    //    string vATNumber = SDOHelper.Read(oPurchaseRecord, "VAT_REG_NUMBER").ToString();

                    //    string website = SDOHelper.Read(oPurchaseRecord, "WWW").ToString();


                    //    if (website.IndexOf(" ") != -1)
                    //        website = website.Remove(website.IndexOf(" "));

                    //    //Regex r = new Regex(@"/((([A-Za-z]{3,9}:(?:\/\/)?)(?:[-;:&=\+\$,\w]+@)?[A-Za-z0-9.-]+|(?:www.|[-;:&=\+\$,\w]+@)[A-Za-z0-9.-]+)((?:\/[\+~%\/.\w-_]*)?\??(?:[-\+=&;%@.\w_]*)#?(?:[\w]*))?)/");
                    //    //// Match the regular expression pattern against a text string.
                    //    //Match m = r.Match(website);
                    //    //while (m.Success)
                    //    //{
                    //    //    //do things with your matching text 
                    //    //    m = m.NextMatch();
                    //    //}

                    //    Account document = new Account
                    //    {
                    //        AccountManagerFullName = accountManagerFullName,
                    //        Blocked = blocked,
                    //        Created = created,
                    //        Type = type,
                    //        BankAccounts = bankAccount,
                    //        CreditLinePurchase = creditLinePurchase,
                    //        Currency = currency,
                    //        GLAccountPurchase = gLAccountPurchase,
                    //        DiscountPurchase = discountPurchase,
                    //        Email = email,
                    //        Fax = fax,
                    //        //EndDate = endDate,
                    //        Name = name,
                    //        PaymentConditionPurchase = paymentConditionPurchase,
                    //        InvoicingMethod = invoicingMethod,
                    //        Website = website,
                    //        IsSupplier = true
                    //    };

                    //    bool createdClient = client.For<Account>().Insert(ref document);

                    } while (oPurchaseRecord.MoveNext());

                    System.IO.File.WriteAllText(@"E:\Get_Suppliers.csv", res);

                    oWS.Disconnect();
                }
                else
                {
                    Console.WriteLine("Failed to connect");
                }
            }
        }
    }
}
