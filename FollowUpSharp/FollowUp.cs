using System;
using System.Collections.Generic;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FollowUpSharp
{
    public class FollowUp
    {
        public delegate void ReportProgressDelegate(int percentage);

        /// <summary>
        /// Runs the accounts gathered from the SQL query
        /// through the SMTP email sending process.
        /// </summary>
        public static void SendFollowUps(string selectedQuery, ReportProgressDelegate mailDelegate, List<string> files = null)
        {
            Query imsQFU = new Query(selectedQuery);
            
            List<IMSEntry>  dbInsureds = imsQFU.GetInsureds(selectedQuery);
            dbInsureds = imsQFU.RemoveDuplicateInsureds(dbInsureds);
            //dbInsureds = ChangeEmails(dbInsureds);
            imsQFU = null;            

            // Make a new sheet that contains the database entries returned for the day
            ExcelWriter sheet = new ExcelWriter();
            sheet.AddResults(dbInsureds);
            sheet.SaveWS();
            sheet = null;

            Outlook.Application olApp = new Outlook.Application();

            for (int i = 0; i < dbInsureds.Count; i++)
            {
                // For targeting the majority OS in the office (Windows 7), moved project to .NET Framework
                // v3.5, and in v3.5, Outlook mailitem objects must be explicitly cast to MailItem since the
                // process in 3.5 returns a "dynamic" object (https://msdn.microsoft.com/en-us/library/dd264736.aspx)
                Outlook.MailItem olMsg = olApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                string htmlBody = $@"", wkfcBody, corriskBody, signature;
                olMsg.To = dbInsureds[i].BrokerEmail;
                //olMsg.CC = dbInsureds[i].UnderwriterEmail;
                olMsg.Subject = $"[CN: {dbInsureds[i].ControlNum}] -- {dbInsureds[i].NamedInsured} -- {dbInsureds[i].EffectiveDate}";

                // If there's no email address on file, don't stop, continue sending out the emails
                if (olMsg.To == null)
                    continue;

                if (selectedQuery.Equals("Quote Follow Ups"))
                {
                    wkfcBody = $@"<!DOCTYPE html>
                    <html>
                    <body><font face=""calibri"" color=""black"" size=""3"">
                    <p>Dear {dbInsureds[i].FirstName},</p>
                    <p>On behalf of WKFC Underwriting Managers, I’m following up on our quote for the account referenced above.
                    If there is anything else you need to bind this account, please let us know. Otherwise, we would be grateful
                    if you could advise us of the status of our quote.</p>
                    <p>In the interest of serving you better in the future, if we have lost this account due to pricing and/or terms,
                    please let us know.</p>
                    <p>Regards,</p>";

                    htmlBody += wkfcBody;
                }
                else if (selectedQuery.Equals("CorRisk Quote Follow Ups"))
                {
                    corriskBody = $@"<!DOCTYPE html>
                    <html>
                    <body><font face=""calibri"" color=""black"" size=""3"">
                    <p>Dear {dbInsureds[i].FirstName},</p>
                    <p>On behalf of CorRisk Solutions, I’m following up on our quote for the account referenced above.
                    If there is anything else you need to bind this account, please let us know. Otherwise, we would be grateful
                    if you could advise us of the status of our quote.</p>
                    <p>In the interest of serving you better in the future, if we have lost this account due to pricing and/or terms,
                    please let us know.</p>
                    <p>Regards,</font></p>";

                    htmlBody += corriskBody;
                }
                signature = FromEmployee(Environment.UserName);
                htmlBody += signature;
                olMsg.HTMLBody = htmlBody;

                if (files != null)
                {
                    foreach (string file in files)
                    {
                        olMsg.Attachments.Add(file);
                    }
                }

                olMsg.Send();
                olMsg = null;

                // ReportDelegateProgress - controls the progress bar incrementation
                if (i == dbInsureds.Count - 1)
                {
                    mailDelegate(100);
                }
                else
                {
                    mailDelegate((i * 100) / dbInsureds.Count - 1);
                }
                Thread.Sleep(500);
            }
            olApp = null;
            dbInsureds = null;
        }

        // This is a debug method, do not uncomment unless debugging!
        // This method, after the database query, will replace all email addresses
        // of the broker with my email address such that the process can be simulated
        // from start to finish without the brokers receiving the emails
        /*
        private static List<IMSEntry> ChangeEmails(List<IMSEntry> accounts)
        {
            foreach (IMSEntry insured in accounts)
            {
                insured.BrokerEmail = "jmarkman@wkfc.com";
            }

            return accounts;
        }
        */
        
        /// <summary>
        /// Expands the user's signature based on the information provided from the environment
        /// variable. 
        /// </summary>
        /// <param name="user">A string returned from Environment.UserName</param>
        /// <returns>A string from the value associated with the UserName key</returns>
        private static string FromEmployee(string user)
        {
            string currentUser = user.ToLower();
            string ta = "";
            Dictionary<string, string> taNames = new Dictionary<string, string>();
            taNames.Add("jmarkman", $@"
                    <p><font face =""CenturyGothic"" color=""#1f497d"" size=""12px""><b>Jonathan Markman</b><br/>
                    TA | Operations<br />
                    <b>WKF&C Underwriting Managers</b><br />
                    (631) 756-3000 x127<br />
                    <a href = mailto:jmarkman@wkfc.com>jmarkman@wkfc.com</a></font><br />
                    <font color=""#00b050"" face=""Arial""><b>Now offering Layers Professional Liability!</b></font><br/>
                    <font size =""1"" face =""CenturyGothic"" color=""#1f497d"">WKFC Underwriting Managers is a series of RSG Underwriting Managers, LLC,;<br/>
                    a Delaware Series Limited Liability Company.In California:< br />
                    RSG Insurance Services, LLC License #0E50879</font></p>
                    </font></body>
                    </html>");

            taNames.Add("stazari", $@"
                    <p><font face =""CenturyGothic"" size=""10px""><b>CorRisk Operations</b><br/>
                    <b>CorRisk Solutions</b><br />
                    <b>O.</b>(631) 756-3000<br />
                    <a href=""https://www.corrisksolutions.com"">www.CorRiskSolutions.com</a><br />                 
                    </p>
                    </font></body>
                    </html>");

            taNames.Add("yzhou", $@"
                    <p><font face =""CenturyGothic"" size=""10px""><b>CorRisk Operations</b><br/>
                    <b>CorRisk Solutions</b><br />
                    <b>O.</b>(631) 756-3000<br />
                    <a href=""https://www.corrisksolutions.com"">www.CorRiskSolutions.com</a><br />                 
                    </p>
                    </font></body>
                    </html>");

            taNames.Add("mpilorge", $@"
                    <p><font face =""CenturyGothic"" size=""10px""><b>CorRisk Operations</b><br/>
                    <b>CorRisk Solutions</b><br />
                    <b>O.</b>(631) 756-3000<br />
                    <a href=""https://www.corrisksolutions.com"">www.CorRiskSolutions.com</a><br />                 
                    </p>
                    </font></body>
                    </html>");

            taNames.Add("tbirnstill", $@"
                    <p><font face =""CenturyGothic"" size=""10px""><b>CorRisk Operations</b><br/>
                    <b>CorRisk Solutions</b><br />
                    <b>O.</b>(631) 756-3000<br />
                    <a href=""https://www.corrisksolutions.com"">www.CorRiskSolutions.com</a><br />                 
                    </p>
                    </font></body>
                    </html>");

            taNames.Add("cmoore", $@"
                    <p><font face =""CenturyGothic"" size=""10px""><b>CorRisk Operations</b><br/>
                    <b>CorRisk Solutions</b><br />
                    <b>O.</b>(631) 756-3000<br />
                    <a href=""https://www.corrisksolutions.com"">www.CorRiskSolutions.com</a><br />                 
                    </p>
                    </font></body>
                    </html>");

            taNames.Add("dbenjamin", $@"
                    <p><font face =""CenturyGothic"" color=""#1f497d"" size=""12px""><b>Darius Benjamin</b><br/>
                    TA | Operations<br />
                    <b>WKF&C Underwriting Managers</b><br />
                    (631) 756-3000 x127<br />
                    <a href = mailto:dbenjamin@wkfc.com>dbenjamin@wkfc.com</a></font><br />
                    <font color=""#00b050"" face=""Arial""><b>Now offering Layers Professional Liability!</b></font><br/>
                    <font size =""1"" face =""CenturyGothic"" color=""#1f497d"">WKFC Underwriting Managers is a series of RSG Underwriting Managers, LLC,;<br/>
                    a Delaware Series Limited Liability Company.In California:< br />
                    RSG Insurance Services, LLC License #0E50879</font></p>
                    </font></body>
                    </html>");

            if (taNames.ContainsKey(currentUser))
            {
                ta = taNames[currentUser];
            }
            return ta;
        }
    }
}
