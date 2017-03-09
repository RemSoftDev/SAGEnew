using System;
using SAGE;
using ConsoleApplication;
using ExactOnline.Client.Sdk.Controllers;
using ExactOnline.Client.Models;

namespace ConsoleApplication
{
    class Get_Quotations
    {

        public Get_Quotations(string companyName)
        {
            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();
            SageDataObject220.WorkSpace oWS;
            SageDataObject220.SopData oSopData;

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
                        //Instantiate Objects
                        oSopData = (SageDataObject220.SopData)oWS.CreateObject("SopData");

                        oSopData.MoveFirst();

                        //Find the first record

                        do
                        {
                            //Populate the fields copying from the existing records
                            if (SDOHelper.Read(oSopData, "QUOTE_STATUS").ToString() != "0")
                            {
                                res = SDOHelper.Read(oSopData, "NAME").ToString() + "  ";
                                Console.WriteLine(res);

                            }




                        } while (oSopData.MoveNext());

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
