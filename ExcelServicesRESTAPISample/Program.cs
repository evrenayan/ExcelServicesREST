namespace ExcelServicesRESTAPISample
{
    using System;
    using System.Collections.Generic;
    using System.Security;
    using System.Runtime.Serialization;
    using System.Runtime.Serialization.Json;
    using Microsoft.SharePoint.Client;
    using System.Net;
    using System.IO;
    using System.Text;
    using System.Xml;
    class Program
    {
        static void Main(string[] args)
        {
            // Get credention information for your SharePoint Online environment
            Console.WriteLine("SharePoint Online username (email):");
            string username = Console.ReadLine();

            SecureString securePassword = new SecureString();
            Console.Write("SharePoint Online password:");

            // Convert password to SecureString
            while (true)
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                if (key.Key == ConsoleKey.Enter)
                    break;
                else if (key.Key == ConsoleKey.Escape)
                    return;
                else if (key.Key == ConsoleKey.Backspace)
                {
                    if (securePassword.Length != 0)
                        securePassword.RemoveAt(securePassword.Length - 1);
                }
                else
                    securePassword.AppendChar(key.KeyChar);
            }

            securePassword.MakeReadOnly();

            // Trying to connect SharePoint Online
            SharePointOnlineCredentials cred = new SharePointOnlineCredentials(username, securePassword);

            Console.WriteLine();
            Console.WriteLine("Trying to connect...");

            // Give parameters for sample excel workbook (with range or table name)
            string url = "https://your-site-url.sharepoint.com/sites/dev/_vti_bin/ExcelRest.aspx/Shared%20Documents/Book1.xlsx/Model/Ranges('TEST1RANGE')?$format=json";

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Credentials = cred;
            req.Headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";
            HttpWebResponse response = (HttpWebResponse)req.GetResponse();

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(RangeResponse));
            RangeResponse rr = serializer.ReadObject(response.GetResponseStream()) as RangeResponse;
            Console.WriteLine("Value from row index 2 and column index 1 : " + rr.rows[2][1].v);
            Console.ReadLine();
        }

    }

    [DataContract]
    public class CellValue
    {
        [DataMember]
        public object v { get; set; }

        [DataMember]
        public object fv { get; set; }
    }

    [DataContract]
    public class RangeResponse
    {
        [DataMember]
        public string name { get; set; }

        [DataMember]
        public CellValue[][] rows { get; set; }
    }
}



