using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExactOnline.Client.Models;
using ExactOnline.Client.Sdk.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApplication
{
    class Get_ChartGLAccounts
    {
        public Get_ChartGLAccounts(string companyName)
        {
            const string clientId = "b4b7919a-f51c-4272-9576-dbcccffb8764";
            const string clientSecret = "eWfc7W7BQXYZ";

            //This can be any url as long as it is identical to the callback url you specified for your app in the App Center.
            var callbackUrl = new Uri("http://websocketclient.local");

            var connector = new Connector(clientId, clientSecret, callbackUrl);
            var client = new ExactOnlineClient(connector.EndPoint, connector.GetAccessToken);


            using (SpreadsheetDocument document = SpreadsheetDocument.Open(@"E:\робочий стіл 2016\Gob Progect\exactonline-api-dotnet-client-master11\exactonline-api-dotnet-client-master\src\ConsoleApplication\Excel\1 Chart of Accounts revised.xlsx", false))
            {
                // open the document read-only

                SharedStringTable sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                string cellValue = null;

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.WorksheetParts.Where(x => x.Uri.OriginalString.Contains("sheet1")).Single();

                foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                {
                    if (sheetData.HasChildren)
                    {
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                cellValue = cell.InnerText;

                                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                {
                                    Console.WriteLine("cell val: " + sharedStringTable.ElementAt(Int32.Parse(cellValue)).InnerText);

                                }
                                else
                                {
                                    Console.WriteLine("else: " + cellValue);
                                }
                            }
                        }
                    }
                }
                document.Close();
            }

            

            try
            {
                //GLAccount document = new GLAccount
                //{
                //    Code = "1001",
                //    Description = "Stock",
                //    TypeDescription = "Current Assets",
                //    BalanceSide = "D",
                //    BalanceType = "B",
                //    Type = 32
                //};

                //bool createdClient = client.For<GLAccount>().Insert(ref document);

                //GLAccount document1 = new GLAccount
                //{
                //    Code = "1220",
                //    Description = "Bank Deposit A/C 19057399",
                //    TypeDescription = "Current Assets",
                //    BalanceSide = "D",
                //    BalanceType = "B",
                //    Type = 12
                //};

                //bool createdClient1 = client.For<GLAccount>().Insert(ref document1);


                //GLAccount document2 = new GLAccount
                //{
                //    Code = "1250",
                //    Description = "Credit Card Account 0793 Paul",
                //    TypeDescription = "Current Assets"
                //};

                //bool createdClient2 = client.For<GLAccount>().Insert(ref document2);

                GLAccount document4 = new GLAccount
                {
                    Code = "1260",
                    Description = "Pauls Cash Expenses",
                    TypeDescription = "Current Liabilities"
                };

                bool createdClient4 = client.For<GLAccount>().Insert(ref document4);

                GLAccount document5 = new GLAccount
                {
                    Code = "2210",
                    Description = "PAYE/PRSI Payable",
                    TypeDescription = "Current Liabilities"
                };

                bool createdClient5 = client.For<GLAccount>().Insert(ref document5);
            }
            catch (Exception ex)
            {
                Console.WriteLine("SDO Generated the Following Error: \n\n" + ex.Message, "Error!");
            }

        }
    }
}
