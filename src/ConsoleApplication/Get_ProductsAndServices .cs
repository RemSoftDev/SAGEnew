using SAGE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Get_ProductsAndServices
    {
        public Get_ProductsAndServices(string companyName)
        {
            //Declare Variables

            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();

            SageDataObject220.WorkSpace oWS;
            SageDataObject220.InvoiceItem oStockPost;

            SageDataObject220.StockTran oStockTran;
            
            SageDataObject220.StockRecord oStockRecord;
            SageDataObject220.StockData oStockData;

            String szDataPath;

            //Instantiate WorkSpace
            oWS = (SageDataObject220.WorkSpace)oSDO.Workspaces.Add("Example");

            //Show select company dialog
            szDataPath = oSDO.SelectCompany(companyName);
            string res = "";

            //Try a connection, will throw an exception if it fails
            try
            {
                //Leaving the username and password blank generates a login dialog
                if (szDataPath != string.Empty)
                {
                    //Connect to the Data Files
                    if (oWS.Connect(szDataPath, "Manager", "", "Demodata"))
                    {
                        oStockTran = (SageDataObject220.StockTran)oWS.CreateObject("StockTran");
                        oStockData = (SageDataObject220.StockData)oWS.CreateObject("StockData");

                        oStockPost = (SageDataObject220.InvoiceItem)oWS.CreateObject("InvoiceItem");

                        oStockRecord = (SageDataObject220.StockRecord)oWS.CreateObject("StockRecord");

                        //Read the first stock code
                        oStockRecord.MoveFirst();
                        oStockPost.MoveFirst();
                        oStockData.MoveFirst();
                        do
                        {
                            //Fill in header information required
                            res = SDOHelper.Read(oStockRecord, "DESCRIPTION") + " ";
                            res += SDOHelper.Read(oStockRecord, "STOCK_CODE") + " ";
                            Console.WriteLine(res);

                        } while (oStockRecord.MoveNext());


                        //Disconnect
                        oWS.Disconnect();
                    }
                    else
                    {
                        Console.WriteLine("Failed to connect");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SDO Generated the Following Error: \n\n" + ex.Message, "Error!");
            }
        }
    }
}
