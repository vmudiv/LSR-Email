using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Net.Mail;
using System.Configuration;
using System.Security.Principal;
using System.Threading;
using System.Globalization;
using Microsoft.Office.Interop.Outlook;

namespace tfstestapi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string strMailBody = string.Empty;
        string strEmailids = string.Empty;
        string strEmailIds = string.Empty;
        public static System.Data.IDataReader QueryKusto(Kusto.Data.Common.ICslQueryProvider queryProvider, string databaseName, string query)
        {
            var clientRequestProperties = new Kusto.Data.Common.ClientRequestProperties();
            clientRequestProperties.ClientRequestId = Guid.NewGuid().ToString();
            clientRequestProperties.SetOption(Kusto.Data.Common.ClientRequestProperties.OptionNoTruncation, true);

            try
            {
                IPrincipal principal = Thread.CurrentPrincipal;
                IIdentity identity = principal == null ? null : principal.Identity;
                string name = identity == null ? "" : identity.Name;
                Thread.CurrentPrincipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                string userName = "v-mudiv@microsoft.com"; // todo
                string[] roles = { "Manager", "Admin" }; // todo
                Thread.CurrentPrincipal = new GenericPrincipal(new GenericIdentity(userName), roles);
                return queryProvider.ExecuteQuery(databaseName, query, clientRequestProperties);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(
                    "Failed invoking query '{0}' on Kusto. Contact kustoops@microsoft.com with clientRequestId={1}. Exception: {2}",
                    query, clientRequestProperties.ClientRequestId, ex.ToString());
                return null;
            }
        }

        public DataTable GetData(string strQuery, DataGridView dtGrid, bool isParent, bool isActive)
        {
            #region


            IPrincipal principal = Thread.CurrentPrincipal;
            IIdentity identity = principal == null ? null : principal.Identity;
            string name = identity == null ? "" : identity.Name;
            Thread.CurrentPrincipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            string userName = "v-suvy@microsoft.com"; // todo
            string[] roles = { "Manager", "Admin" }; // todo
            Thread.CurrentPrincipal = new GenericPrincipal(new GenericIdentity(userName), roles);

            #endregion

            string strMbody = string.Empty;
            var client = Kusto.Data.Net.Client.KustoClientFactory.CreateCslQueryProvider("https://icmcluster.kusto.windows.net/IcmDataWarehouse;Fed=true");

            var kcsb = string.Empty;

            var countReader = client.ExecuteQuery(strQuery);
            DataTable dt = new DataTable();
            if (countReader != null)
            {
                dt.Load(countReader);
            }
            return dt;
        }
        static string strTFSEmail = string.Empty;
        public void sendMail(string strFileName, string strMailbody, string strToaddress)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app = new
                           Microsoft.Office.Interop.Outlook.Application();
                MailItem item = app.CreateItem((OlItemType.olMailItem));
                item.BodyFormat = OlBodyFormat.olFormatHTML;

                item.Subject = "Livesite Retrospective" + DateTime.Now;
                item.To = strToaddress;
                item.CC = ConfigurationManager.AppSettings["MailCC"].ToString();
                item.HTMLBody = strMailbody;

                item.Display();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
        static string strMailHeadrer = string.Empty;
        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;

            string strMailBody = string.Empty;
            string strToemailids = string.Empty;
            string strQuery = string.Empty;
            string strPath = Environment.CurrentDirectory + "\\Appcenter.txt";
            strQuery = System.IO.File.ReadAllText(strPath);
            strMailHeadrer += "<html><head><style>table {empty-cells:show; }</style></head><body>";
            strMailHeadrer += "<font color='black'><br>Hello Everyone,<br><br> " + " Listed below are the incidents which are yet to be reviewed in the Livesite Retrospective. A follow ";
            //strMailHeadrer += "<br>";
            strMailHeadrer += "up email with the final list of incidents in <b>Ready for Review </b> to be reviewed this week will be sent at ";
            strMailHeadrer += "<span style='background-color:#FFFF00'> 06:00 PM PST Tuesday (";
            List<string> lstTo = new List<string>();
            int daysUntilTuesday = ((int)DayOfWeek.Tuesday - (int)today.DayOfWeek + 7) % 7;
            DateTime nextTuesday = today.AddDays(daysUntilTuesday);
            DateTime currentmonday = nextTuesday.AddDays(-1);
            DateTime lastmonday = nextTuesday.AddDays(-8);
            DateTime lastlastmonday = lastmonday.AddDays(-7);
            DateTime lastsunday = lastmonday.AddDays(-1);
            DateTime lastLastsunday = lastmonday.AddDays(-8);
            strMailHeadrer += nextTuesday.Month.ToString("00") + "/" + nextTuesday.Day.ToString("00") + ").</span>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "Owners of the below incidents without postmortems are expected to be ready with the " + "<a href='https://msmobilecenter.visualstudio.com/Mobile-Center/_wiki/wikis/Mobile-Center.wiki?wikiVersion=GBwikiMaster&pagePath=%2FHome%2FLive%20Site%2FLive%20Site%20Management%20Tools%2FPost%20Mortem%20Report&pageId=405'>postmortem document </a>";
            strMailHeadrer += " as per the guidelines available  ";
            strMailHeadrer += "<a href='https://msmobilecenter.visualstudio.com/Mobile-Center/_wiki/wikis/Mobile-Center.wiki?wikiVersion=GBwikiMaster&pagePath=%2FHome%2FLive%20Site%2FLive%20Site%20Management%20Tools%2FPost%20Mortem%20Report&pageId=405'> HERE</a>";
            strMailHeadrer += " before 05:00 PM PST Tuesday (" + nextTuesday.Month.ToString("00") + "/" + nextTuesday.Day.ToString("00") + "). I will follow up with incident owners so they are aware of the incidents which needs postmortem.";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<b>If any of the below incidents were not customer impacting, please reduce the severity to 3 by providing the justification in the IcM ticket so it can be removed from the list. If multiple incidents in the below list have the same root cause, please link them as child items to the first occurrence of the issue and create the postmortem for the parent incident. Please let me know if you have any questions.</b>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<b><u>App Center</u></b>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += " <table border=1><tr><th>Id</th><th>Severity</th><th>Status</th><th>Title</th><th>CreateDate</th><th>TeamName</th><th>Owner</th><th>Postmortem</th></tr>";
            DataTable dt1 = GetData(strQuery, new DataGridView(), false, true);
            foreach (DataRow dr in dt1.Rows)
            {
                DateTime dtCreateDate = Convert.ToDateTime(dr[4].ToString());
                if (dtCreateDate <= currentmonday)
                {
                    string stricmURL = "https://icm.ad.msft.net/imp/v3/incidents/details/" + dr[0].ToString() + "/home";
                    string strPoststatus = string.Empty;
                    if (!string.IsNullOrEmpty(dr[7].ToString()))
                    {
                        string strposturl = string.Empty;
                        strposturl += "https://icm.ad.msft.net/imp/v3/incidents/postmortem/" + dr[7].ToString();
                        strPoststatus += "<a href='" + strposturl + "'>" + dr[7].ToString() + "</a>";
                    }

                    if (dtCreateDate > lastmonday)
                    {
                        strMailHeadrer += "<tr>";
                    }
                    else if (dtCreateDate >= lastlastmonday && dtCreateDate <= lastmonday)
                    {
                        strMailHeadrer += "<tr bgcolor='#FFE599'>";
                    }
                    else if (dtCreateDate <= lastlastmonday)
                    {
                        strMailHeadrer += "<tr bgcolor='#F7CAAC'>";
                    }
                    if (!string.IsNullOrEmpty(dr[8].ToString()))
                    {
                        strPoststatus = strPoststatus + " - " + dr[8].ToString();
                    }
                    if (!string.IsNullOrEmpty(dr[6].ToString()) && !string.IsNullOrEmpty(strPoststatus))
                    {
                        strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + "</a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + dr[6].ToString() + "@microsoft.com" + "</td><td>" + strPoststatus + "</td></tr>";
                    }
                    else if (string.IsNullOrEmpty(dr[6].ToString()))
                    {
                        strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + "&nbsp;" + "</td><td>" + strPoststatus + "</td></tr>";
                    }
                    else if (string.IsNullOrEmpty(strPoststatus))
                    {
                        strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + dr[6].ToString() + "@microsoft.com" + "</td><td>" + "&nbsp;" + "</td></tr>";
                    }
                    else if (string.IsNullOrEmpty(dr[6].ToString()) && string.IsNullOrEmpty(strPoststatus))
                    {
                        strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + "&nbsp;" + "</td><td>" + "&nbsp;" + "</td></tr>";
                    }
                    string email = dr[6].ToString() + "@microsoft.com";
                    if (!string.IsNullOrEmpty(dr[6].ToString()) && !lstTo.Contains(email))
                    {
                        lstTo.Add(email);
                    }
                    strMailHeadrer += "</tr>";
                }
            }
            strMailHeadrer += "</table>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<b><u>Notification Hubs</u></b>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += " <table border=1><tr><th>Id</th><th>Severity</th><th>Status</th><th>Title</th><th>CreateDate</th><th>TeamName</th><th>Owner</th><th>Postmortem</th></tr>";
            string strQuery1 = string.Empty;
            string strPath1 = Environment.CurrentDirectory + "\\AzNHub.txt";
            strQuery1 = System.IO.File.ReadAllText(strPath1);
            DataTable dt2 = GetData(strQuery1, new DataGridView(), false, true);
            foreach (DataRow dr in dt2.Rows)
            {
                DateTime dtCreateDate = Convert.ToDateTime(dr[4].ToString());
                if (dtCreateDate < currentmonday)
                {
                    string strStatus = dr[2].ToString();
                    string strPIRId = dr[7].ToString();
                    string strPIRstatus = dr[8].ToString();

                    if ((strStatus != "RESOLVED") || (strStatus == "RESOLVED" && !string.IsNullOrEmpty(strPIRId) && strPIRstatus != "Completed"))
                    {

                        string stricmURL = "https://icm.ad.msft.net/imp/v3/incidents/details/" + dr[0].ToString() + "/home";
                        string strPoststatus = string.Empty;
                        if (!string.IsNullOrEmpty(dr[7].ToString()))
                        {
                            string strposturl = string.Empty;
                            strposturl += "https://icm.ad.msft.net/imp/v3/incidents/postmortem/" + dr[7].ToString();
                            strPoststatus += "<a href='" + strposturl + "'>" + dr[7].ToString() + "</a>";
                        }

                        if (dtCreateDate > lastmonday)
                        {
                            strMailHeadrer += "<tr>";
                        }
                        else if (dtCreateDate >= lastlastmonday && dtCreateDate <= lastmonday)
                        {
                            strMailHeadrer += "<tr bgcolor='#FFE599'>";
                        }
                        else if (dtCreateDate <= lastlastmonday)
                        {
                            strMailHeadrer += "<tr bgcolor='#F7CAAC'>";
                        }
                        if (!string.IsNullOrEmpty(dr[8].ToString()))
                        {
                            strPoststatus = strPoststatus + " - " + dr[8].ToString();
                        }
                        if (!string.IsNullOrEmpty(dr[6].ToString()) && !string.IsNullOrEmpty(strPoststatus))
                        {
                            strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + "</a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + dr[6].ToString() + "@microsoft.com" + "</td><td>" + strPoststatus + "</td></tr>";
                        }
                        else if (string.IsNullOrEmpty(dr[6].ToString()))
                        {
                            strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + "&nbsp;" + "</td><td>" + strPoststatus + "</td></tr>";
                        }
                        else if (string.IsNullOrEmpty(strPoststatus))
                        {
                            strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + dr[6].ToString() + "@microsoft.com" + "</td><td>" + "&nbsp;" + "</td></tr>";
                        }
                        else if (string.IsNullOrEmpty(dr[6].ToString()) && string.IsNullOrEmpty(strPoststatus))
                        {
                            strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[2].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + dr[4].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + "&nbsp;" + "</td><td>" + "&nbsp;" + "</td></tr>";
                        }

                        // strMailHeadrer += "</tr>";
                        string email = dr[6].ToString() + "@microsoft.com";
                        if (!string.IsNullOrEmpty(dr[6].ToString()) && !lstTo.Contains(email))
                        {
                            lstTo.Add(email);
                        }
                    }
                }

            }
            strMailHeadrer += "</table>";
            //Severity Bump Down data
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<b><u>Severity Bumped-Down Incidents</u></b>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "This section highlights the IcM Incidents where the severity was bumped-down and the justification for the same.";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += " <table border=1><tr><th>Id</th><th>ChangeDate_PST</th><th>ChangedBy</th><th>Severity</th><th>Reason</th><th>TeamName</th><th>Title</th><th>HowFixed</th><th>Status</th></tr>";
            string strQuery2 = string.Empty;
            string strPath2 = Environment.CurrentDirectory + "\\SevDowngraded.txt";
            strQuery2 = System.IO.File.ReadAllText(strPath2);
            DataTable dt3 = GetData(strQuery2, new DataGridView(), false, true);
            foreach (DataRow dr in dt3.Rows)
            {
                DateTime dtCreateDate = Convert.ToDateTime(dr[2].ToString());

                if (dtCreateDate < currentmonday)
                {
                    string stricmURL = "https://icm.ad.msft.net/imp/v3/incidents/details/" + dr[0].ToString() + "/home";
                    string reasonRaw = (dr[4].ToString()).Replace("Severity change from 2 to 3.", "");
                    string reason = reasonRaw.Replace("Reason", "");
                    if (dtCreateDate > lastmonday)
                    {
                        strMailHeadrer += "<tr>";
                        strMailHeadrer += "<td><a href='" + stricmURL + "'>" + dr[0].ToString() + " </a></td><td>" + dr[1].ToString() + "</td><td>" + dr[5].ToString() + "</td><td>" + dr[3].ToString() + "</td><td>" + reason + "</td><td>" + dr[6].ToString() + "</td><td>" + dr[7].ToString() + "</td><td>" + dr[8].ToString() + "</td><td>" + dr[9].ToString() + "</td></tr>";
                    }
                }
            }
            strMailHeadrer += "</table>";

            string toemail = string.Empty;
            foreach (string str in lstTo)
            {
                toemail += str + ";";
            }

            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<table border='1'><tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>Incidents created within the last one week</td><td>" + lastmonday.Month + "/" + lastmonday.Day + " or later" + "</td></tr>";
            strMailHeadrer += "<tr><td bgcolor='#FFE599' >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>Incidents older than a week but newer than two weeks</td><td>" + lastlastmonday.Month + "/" + lastlastmonday.Day + " to " + lastsunday.Month + "/" + lastsunday.Day + " </td></tr>";
            strMailHeadrer += "<tr ><td bgcolor='F7CAAC' > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>Incidents older than two weeks</td><td>" + lastLastsunday.Month + "/" + lastLastsunday.Day + " or before" + " </td></tr>";
            strMailHeadrer += "</table>";
            strMailHeadrer += "<br>";
            strMailHeadrer += "<br>";
            writeHtmlOutput(strMailHeadrer);
            //below code to send email
            sendMail("", strMailHeadrer, toemail);
            this.Close();
        }

        void writeHtmlOutput(string htmlOutput)
        {
            string outputPath = ConfigurationManager.AppSettings["OutputPath"].ToString();
            string temp = DateTime.Now.ToString("YYYYMMddhhmmssfff", CultureInfo.InvariantCulture);
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(outputPath + @"\HtmlOutput" + DateTime.Now.ToString("yyyyMMddhhmmssfff") + ".html", false))
            {
                file.Write(htmlOutput);
            }
        }

    }

}

