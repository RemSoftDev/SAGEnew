using ExactOnline.Client.Models;
using ExactOnline.Client.Sdk.Controllers;
using SAGE;
using System;
using System.Diagnostics;
using System.Linq;

namespace ConsoleApplication
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
            string companyName = @"C:\ProgramData\Sage\Accounts\2016\";


   //         // These are the authorisation properties of your app.
   //         // You can find the values in the App Center when you are maintaining the app.
   //         const string clientId = "b4b7919a-f51c-4272-9576-dbcccffb8764";
			//const string clientSecret = "eWfc7W7BQXYZ";

			//// This can be any url as long as it is identical to the callback url you specified for your app in the App Center.
			//var callbackUrl = new Uri("http://websocketclient.local"); 

   //         var connector = new Connector(clientId, clientSecret, callbackUrl);
			//var client = new ExactOnlineClient(connector.EndPoint, connector.GetAccessToken);

   //         Account document = new Account
   //         {
   //             Name = "rererererer",
   //             City = "trololo"
   //         };

   //         bool created = client.For<Account>().Insert(ref document);

   //         // Get the Code and Name of a random account in the administration
   //         var fields = new[] { "Code", "Name" };
			//var account = client.For<Account>().Top(1).Select(fields).Get().FirstOrDefault();

   //         Console.WriteLine("Account {0} - {1}", account.Code.TrimStart(), account.Name);

            // new Get_Transactions(companyName);
            //new Get_Suppliers(companyName);
            new Get_ChartGLAccounts(companyName);
            Console.ReadLine();
        }
	}
}
