using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Net;
using Newtonsoft.Json;

namespace UsageDashboardHelper
{
    public class Data
    {
        public string gid { get; set; }
        public string resource_type { get; set; }
        public string name { get; set; }
    }

    public class GenericData
    {
        public List<Data> data { get; set; }
    }

    class Program
    {
        public static string dbPath = "Provider = SQLOLEDB; Initial Catalog = UsageDashboard; Data Source = mssql01; Integrated Security = SSPI;";

        private static string GetThreeUsageReport(string sRN, string sRCN, int iRID, string sUserName, string sUserPassword, int sAuthCompanyId)
        {
            string sResult = "n/a";

            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            MachPanelWebService.MPAuthenticationHeader header = new MachPanelWebService.MPAuthenticationHeader();

            header.UserName = sUserName;
            header.UserPassword = sUserPassword;

            header.AuthenticationToken = "";
            header.CompanyId = sAuthCompanyId;

            MachPanelWebService.MachPanelService svc = new MachPanelWebService.MachPanelService();
            svc.Url = "https://controlpanel.gowingu.net/webservices/machpanelservice.asmx";
            svc.MPAuthenticationHeaderValue = header;

            string sResellerName = sRN;
            string sResellerCompanyName = sRCN;
            int iResellerId = iRID;

            // MachPanel data
            MachPanelWebService.ResponseArguments ra = svc.Authenticate();
            //MachPanelWebService.Customer[] allCustomers = svc.GetAllCustomers();
            //MachPanelWebService.SubscriptionInfo[] allSubs = svc.GetAllSubscriptions();
            //MachPanelWebService.PaymentGroup[] AllPg = svc.GetAllPaymentGroups();

            Console.WriteLine("Name: " + sResellerName + "; Company Name: " + sResellerCompanyName + " [Id:" + iResellerId.ToString() +"]");

            // list all usage report
            string sAllUsagePrintReportData = "";
            try
            {
                MachPanelWebService.ReportingCriteria RC = new MachPanelWebService.ReportingCriteria();
                MachPanelWebService.LyncUserUsageReport[] lyncuserUsageReportList = svc.GetLyncUserUsageReport(RC);
                if (lyncuserUsageReportList.Length > 0)
                {
                    int iCount = 0;

                    sAllUsagePrintReportData += "ArchivingPolicy,ClientPolicy,ClientVersionPolicy,CompanyId,CompanyName,ConferencingPolicy,CustomerID,CustomerName,CustomerNumber,DateCreated,DialPlan,ExternalAccessPolicy,IsChatEnabled,LocationPolicy,LyncUser,MicrosoftSPLAType,MobilityPolicy,OrganizationName,Owner,PersistentChatPolicy,PhoneNumber,PinPolicy,ResellerId,SoldPackage,TelephonyOption,VoiceMailPolicy,VoicePolicy\r\n";

                    foreach (MachPanelWebService.LyncUserUsageReport lur in lyncuserUsageReportList)
                    {
                        if (lur.ResellerId == iResellerId)
                        {
                            sAllUsagePrintReportData += "\"" + lur.ArchivingPolicy + "\"," + "\"" + lur.ClientPolicy + "\"," + "\"" + lur.ClientVersionPolicy + "\"," + lur.CompanyId + "," + "\"" + lur.CompanyName + "\"," + "\"" + lur.ConferencingPolicy + "\"," + lur.
                            CustomerID + "," + "\"" + lur.CustomerName + "\"," + "\"" + lur.CustomerNumber + "\"," + "\"" + lur.DateCreated + "\"," + "\"" + lur.DialPlan + "\"," + "\"" + lur.ExternalAccessPolicy + "\"," + "" + lur.
                            IsChatEnabled + "," + "\"" + lur.LocationPolicy + "\"," + "\"" + lur.LyncUser + "\"," + "\"" + lur.MicrosoftSPLAType + "\"," + "\"" + lur.MobilityPolicy + "\"," + "\"" + lur.OrganizationName + "\"," + "\"" + lur.
                            Owner + "\"," + "\"" + lur.PersistentChatPolicy + "\"," + "\"" + lur.PhoneNumber + "\"," + "\"" + lur.PinPolicy + "\"," + lur.ResellerId + "," + "\"" + lur.SoldPackage + "\"," + "\"" + lur.TelephonyOption + "\"," + "\"" + lur.
                            VoiceMailPolicy + "\"," + "\"" + lur.VoicePolicy + "\r\n";
                            iCount++;
                        }
                    }

                    Console.WriteLine(iCount.ToString() + " user reports found.");
                }

                // write csv file
                try
                {
                    if (sAllUsagePrintReportData != "")
                    {
                        string sFileName = sResellerCompanyName + "_" + sResellerName + "_";
                        sFileName += DateTime.Now.Day.ToString().PadLeft(2, '0');
                        sFileName += DateTime.Now.Month.ToString().PadLeft(2, '0');
                        sFileName += DateTime.Now.Year.ToString().PadLeft(4, '0');
                        sFileName += ".csv";
                        string sFileNameLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + sFileName;

                        try
                        {
                            File.WriteAllText(sFileNameLocation, sAllUsagePrintReportData);
                            Console.WriteLine("CSV file location: " + sFileNameLocation);
                            sResult = sFileNameLocation;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }
                    else
                    {
                        Console.WriteLine("No usage report found at the moment.");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    sResult = "n/a";
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                sResult = "n/a";
            }

            Console.WriteLine("");

            return sResult;
        }

        private static int GetRackpeopleUsersNumber(string sRN, string sRCN, int iRID, string sUserName, string sUserPassword, int sAuthCompanyId)
        {
            int iResult = 0;

            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            MachPanelWebService.MPAuthenticationHeader header = new MachPanelWebService.MPAuthenticationHeader();

            header.UserName = sUserName;
            header.UserPassword = sUserPassword;

            header.AuthenticationToken = "";
            header.CompanyId = sAuthCompanyId;

            MachPanelWebService.MachPanelService svc = new MachPanelWebService.MachPanelService();
            svc.Url = "https://controlpanel.gowingu.net/webservices/machpanelservice.asmx";
            svc.MPAuthenticationHeaderValue = header;

            string sResellerName = sRN;
            string sResellerCompanyName = sRCN;
            int iResellerId = iRID;

            // MachPanel data
            MachPanelWebService.ResponseArguments ra = svc.Authenticate();
            //MachPanelWebService.Customer[] allCustomers = svc.GetAllCustomers();
            //MachPanelWebService.SubscriptionInfo[] allSubs = svc.GetAllSubscriptions();
            //MachPanelWebService.PaymentGroup[] AllPg = svc.GetAllPaymentGroups();

            Console.WriteLine("Name: " + sResellerName + "; Company Name: " + sResellerCompanyName + " [Id:" + iResellerId.ToString() + "]");

            // list all usage report
            string sAllUsagePrintReportData = "";
            try
            {
                MachPanelWebService.ReportingCriteria RC = new MachPanelWebService.ReportingCriteria();
                MachPanelWebService.LyncUserUsageReport[] lyncuserUsageReportList = svc.GetLyncUserUsageReport(RC);
                if (lyncuserUsageReportList.Length > 0)
                {
                    foreach (MachPanelWebService.LyncUserUsageReport lur in lyncuserUsageReportList)
                    {
                        if (lur.ResellerId == iResellerId)
                        {
                            sAllUsagePrintReportData += "\"" + lur.ArchivingPolicy + "\"," + "\"" + lur.ClientPolicy + "\"," + "\"" + lur.ClientVersionPolicy + "\"," + lur.CompanyId + "," + "\"" + lur.CompanyName + "\"," + "\"" + lur.ConferencingPolicy + "\"," + lur.
                            CustomerID + "," + "\"" + lur.CustomerName + "\"," + "\"" + lur.CustomerNumber + "\"," + "\"" + lur.DateCreated + "\"," + "\"" + lur.DialPlan + "\"," + "\"" + lur.ExternalAccessPolicy + "\"," + "" + lur.
                            IsChatEnabled + "," + "\"" + lur.LocationPolicy + "\"," + "\"" + lur.LyncUser + "\"," + "\"" + lur.MicrosoftSPLAType + "\"," + "\"" + lur.MobilityPolicy + "\"," + "\"" + lur.OrganizationName + "\"," + "\"" + lur.
                            Owner + "\"," + "\"" + lur.PersistentChatPolicy + "\"," + "\"" + lur.PhoneNumber + "\"," + "\"" + lur.PinPolicy + "\"," + lur.ResellerId + "," + "\"" + lur.SoldPackage + "\"," + "\"" + lur.TelephonyOption + "\"," + "\"" + lur.
                            VoiceMailPolicy + "\"," + "\"" + lur.VoicePolicy + "\r\n";
                            iResult++;
                        }
                    }

                    Console.WriteLine(iResult.ToString() + " SfB users found.");
                }

            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            Console.WriteLine("");

            return iResult;
        }

        private static string GetNewAsanaProjects(string sAsanaAppToken)
        {
            List<string> sAsanaWorkspaces = new List<string>();
            string sAsanaThreeProjectsGid = "";
            string sAsanaS4BProjectsGid = "";

            try
            {
                var webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/workspaces") as HttpWebRequest;
                if (webRequestAsana != null)
                {
                    webRequestAsana.Method = "GET";
                    webRequestAsana.Accept = "application/json";
                    webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                    using (var s = webRequestAsana.GetResponse().GetResponseStream())
                    {
                        using (var sr = new StreamReader(s))
                        {
                            var AsanaMeJson = sr.ReadToEnd();
                            var WorkspacesAsana = JsonConvert.DeserializeObject<GenericData>(AsanaMeJson);
                            for (int i = 0; i < WorkspacesAsana.data.Count; i++)
                            {
                                if (WorkspacesAsana.data[i].gid != null)
                                {
                                    sAsanaWorkspaces.Add(WorkspacesAsana.data[i].gid);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }

            if (sAsanaWorkspaces.Count > 0)
            {
                foreach (string sAsanaWorkspace in sAsanaWorkspaces)
                {
                    try
                    {
                        var webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/organizations/" + sAsanaWorkspace + "/teams") as HttpWebRequest;
                        if (webRequestAsana != null)
                        {
                            webRequestAsana.Method = "GET";
                            webRequestAsana.Accept = "application/json";
                            webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                            using (var s = webRequestAsana.GetResponse().GetResponseStream())
                            {
                                using (var sr = new StreamReader(s))
                                {
                                    var AsanaMeJson = sr.ReadToEnd();
                                    var TeamsAsana = JsonConvert.DeserializeObject<GenericData>(AsanaMeJson);
                                    for (int i = 0; i < TeamsAsana.data.Count; i++)
                                    {
                                        if (TeamsAsana.data[i].gid != null)
                                        {
                                            if (TeamsAsana.data[i].name == "RP S4B & Teams Onboarding Projects")
                                            {
                                                sAsanaS4BProjectsGid = TeamsAsana.data[i].gid;
                                            }

                                            if (TeamsAsana.data[i].name == "HI3G")
                                            {
                                                sAsanaThreeProjectsGid = TeamsAsana.data[i].gid;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.ToString();
                    }
                }
            }

            int iThreeProjects = 0;
            int iS4BProjects = 0;

            if (sAsanaS4BProjectsGid != "")
            {
                try
                {
                    var webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/teams/" + sAsanaS4BProjectsGid + "/projects") as HttpWebRequest;
                    if (webRequestAsana != null)
                    {
                        webRequestAsana.Method = "GET";
                        webRequestAsana.Accept = "application/json";
                        webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                        using (var s = webRequestAsana.GetResponse().GetResponseStream())
                        {
                            using (var sr = new StreamReader(s))
                            {
                                var AsanaMeJson = sr.ReadToEnd();
                                var S4BProjectsAsana = JsonConvert.DeserializeObject<GenericData>(AsanaMeJson);
                                iS4BProjects = S4BProjectsAsana.data.Count;

                                for (int i = 0; i < S4BProjectsAsana.data.Count; i++)
                                {
                                    if (S4BProjectsAsana.data[i].gid != null)
                                    {
                                        // check if project is active
                                        try
                                        {
                                            webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/projects/" + S4BProjectsAsana.data[i].gid) as HttpWebRequest;
                                            if (webRequestAsana != null)
                                            {
                                                webRequestAsana.Method = "GET";
                                                webRequestAsana.Accept = "application/json";
                                                webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                                                using (var s2 = webRequestAsana.GetResponse().GetResponseStream())
                                                {
                                                    using (var sr2 = new StreamReader(s2))
                                                    {
                                                        AsanaMeJson = sr2.ReadToEnd();
                                                        if (AsanaMeJson.IndexOf("\"archived\": true") != -1)
                                                        {
                                                            iS4BProjects--;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            ex.ToString();
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }

            if (sAsanaThreeProjectsGid != "")
            {
                try
                {
                    var webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/teams/" + sAsanaThreeProjectsGid + "/projects") as HttpWebRequest;
                    if (webRequestAsana != null)
                    {
                        webRequestAsana.Method = "GET";
                        webRequestAsana.Accept = "application/json";
                        webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                        using (var s = webRequestAsana.GetResponse().GetResponseStream())
                        {
                            using (var sr = new StreamReader(s))
                            {
                                var AsanaMeJson = sr.ReadToEnd();
                                var TeamsProjectsAsana = JsonConvert.DeserializeObject<GenericData>(AsanaMeJson);
                                iThreeProjects = TeamsProjectsAsana.data.Count;

                                for (int i = 0; i < TeamsProjectsAsana.data.Count; i++)
                                {
                                    if (TeamsProjectsAsana.data[i].gid != null)
                                    {
                                        // check if project is active
                                        try
                                        {
                                            webRequestAsana = WebRequest.Create("https://app.asana.com/api/1.0/projects/" + TeamsProjectsAsana.data[i].gid) as HttpWebRequest;
                                            if (webRequestAsana != null)
                                            {
                                                webRequestAsana.Method = "GET";
                                                webRequestAsana.Accept = "application/json";
                                                webRequestAsana.Headers.Add("Authorization", "Bearer " + sAsanaAppToken);

                                                using (var s2 = webRequestAsana.GetResponse().GetResponseStream())
                                                {
                                                    using (var sr2 = new StreamReader(s2))
                                                    {
                                                        AsanaMeJson = sr2.ReadToEnd();
                                                        if (AsanaMeJson.IndexOf("\"archived\":true") != -1)
                                                        {
                                                            iThreeProjects--;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            ex.ToString();
                                        }
                                    }
                                }
                                
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }

            Console.WriteLine("Asana 3 team projects count: " + iThreeProjects.ToString());
            Console.WriteLine("SfB team projects count: " + iS4BProjects.ToString());

            return iS4BProjects.ToString() + "," + iThreeProjects.ToString();
        }

        public static string InsertUpdateDatabase(string SQL, string dbPath)
        {
            string DatabaseConnection = dbPath;
            string sResult = "DBOK";

            try
            {
                // Database Object instancing here
                OleDbConnection DatabaseFile;
                OleDbCommand OleCommand;

                // Open database connection
                DatabaseFile = new OleDbConnection(@DatabaseConnection);
                DatabaseFile.Open();

                OleCommand = new OleDbCommand(SQL, DatabaseFile);
                OleCommand.ExecuteNonQuery();

                // Close Connection
                DatabaseFile.Close();
            }
            catch (Exception Ex)
            {
                Ex.ToString();
                sResult = SQL + "   ::   " + Ex.ToString();
                return sResult;
            }

            return sResult;
        }

        static void Main(string[] args)
        {
            // Get Usage Report for Ronni Hansen 3
            string sAttFile = GetThreeUsageReport("Ronni Hansen", "3", 79, "provider@gowingu.net", "h6ECyLzvzQ&3sm", 79);

            if (sAttFile != "n/a")
            {
                DateTime dtCurrentTime = DateTime.Now;
                if ((dtCurrentTime.Day == 1) && (dtCurrentTime.Hour == 7))
                {
                    try
                    {
                        System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                        string[] sRecepientEmailArray = "sa@rackpeople.dk,mz@rackpeople.dk".Split(',');
                        foreach (string sRecepient in sRecepientEmailArray)
                        {
                            message.To.Add(sRecepient);
                        }
                        message.Subject = "Monthly Usage Report for customer 3 (Ronni Hansen)";
                        message.IsBodyHtml = false;
                        message.BodyEncoding = System.Text.Encoding.UTF8;
                        message.Body = "Report attached.";

                        System.Net.Mail.Attachment msatt = new System.Net.Mail.Attachment(sAttFile);
                        message.Attachments.Add(msatt);

                        System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("relay.rackpeople.com");
                        message.From = new System.Net.Mail.MailAddress("webmaster@rackpeople.com");
                        smtp.Send(message);

                        Console.WriteLine("Monthly Usage Report for customer 3 (Ronni Hansen) sent to: " + "sa@rackpeople.dk,mz@rackpeople.dk");
                    }
                    catch (Exception ExpEM)
                    {
                        Console.WriteLine(ExpEM.ToString());
                    }

                    // delete old files
                    try
                    {
                        DirectoryInfo d = new DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
                        FileInfo[] Files = d.GetFiles("*.csv");
                        foreach (FileInfo file in Files)
                        {
                            if (file.CreationTime <= DateTime.Now.AddDays(-1))
                            {
                                file.Delete();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.ToString();
                    }
                }
            }

            // SfB users number for Rackpeople ApS
            int iS4BCount = GetRackpeopleUsersNumber("RackPeople ApS", "RackPeople", 87, "sales@rackpeople.dk", "Tu6Y8AH_", 87);

            // Asana numbers
            string sAsanaProjectsNumbers = GetNewAsanaProjects("0/ec8084b64f058d6f849262b071def980");
            int iS4BProjects = Convert.ToInt32(sAsanaProjectsNumbers.Split(',')[0]);
            int iThreeProjects = Convert.ToInt32(sAsanaProjectsNumbers.Split(',')[1]);

            string dtCurrent = DateTime.Now.Year.ToString().PadLeft(4, '0') + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "T" + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0');

            string sSqlQuery = "INSERT INTO [dbo].[SystemCounts] ([LogDate], [SfBUsers], [TeamsUsers], [HI3GActiveProjects], [RPs4BActiveProjects]) ";
            sSqlQuery += "VALUES ('" + dtCurrent + "', " + iS4BCount.ToString() + ", -1, " + iThreeProjects.ToString() + ", " + iS4BProjects + ")";
            string sResult = InsertUpdateDatabase(sSqlQuery, dbPath);

            Console.WriteLine("");

            if (sResult == "DBOK")
            {
                Console.WriteLine("Values stored in UsageDashboard database.");
            }
            else
            {
                Console.WriteLine(sResult);
            }

            //Console.ReadKey();
        }
    }
}
