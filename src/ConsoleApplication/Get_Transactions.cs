using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAGE
{
    public class Get_Transactions
    {
        public Get_Transactions( string companyName)
        {
            //Declare Variables
            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();
            SageDataObject220.WorkSpace oWS;
            SageDataObject220.StockRecord oStockRecord;
            SageDataObject220.StockTran oStockTran;

            String szDataPath;

            //Instantiate WorkSpace
            oWS = (SageDataObject220.WorkSpace)oSDO.Workspaces.Add("Example");

            //Show select company dialog
            szDataPath = oSDO.SelectCompany(companyName);

            //Try a connection, will throw an exception if it fails
            try

            {
                //Leaving the username and password blank generates a login dialog
                oWS.Connect(szDataPath, "Manager", "", "SDO EXAMPLE");
                
                //Instantiate Stock Record Object
                oStockRecord = (SageDataObject220.StockRecord)oWS.CreateObject("StockRecord");
                var count = oStockRecord.Count;

                //Read the first stock record
                oStockRecord.MoveFirst();

                //Read the first transaction
                oStockTran = (SageDataObject220.StockTran)oStockRecord.Link;
                oStockTran.MoveFirst();
                var isWhile = true;

                //Loop through the transactions
                do
                {
                    //You need to call read before you can return the data
                    oStockTran.Read(oStockTran.RecordNumber);

                    //Display the description
                    Console.WriteLine("Description:  " + (string)SDOHelper.Read(oStockTran, "DETAILS"));
                    isWhile = oStockTran.MoveNext();
                        } while (isWhile);

                //Disconnect
                oWS.Disconnect();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SDO Generated the Following Error: \n\n" + ex.Message, "Error!");
            }
        }
    }
}
