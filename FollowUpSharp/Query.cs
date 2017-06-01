using FollowUpSharp.Exceptions;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;
using System;
using System.IO;
using System.Configuration;

namespace FollowUpSharp
{
    public class Query
    {
        private SqlConnection dbConn;
        private string errorLogPath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\FollowUpSharp\";

        public Query(string selectedQuery)
        {
            // ActiveDirectory connection, will be replaced on deployment
            dbConn = new SqlConnection(GetConnection(selectedQuery));
        }

        /// <summary>
        /// Fetches the list of active control numbers based on the QFU query.
        /// </summary>
        /// <returns>Returns a string list of control numbers.</returns>
        public List<IMSEntry> GetInsureds(string selectedQuery)
        {
            string query;
            List<IMSEntry> insuredStorage = new List<IMSEntry>();
            try
            {
                dbConn.Open();
                // TODO: See if I can beg someone at MGA to let me have these as stored procedures
                switch (selectedQuery.ToLower())
                {
                    case "quote follow ups":
                        query = @"";
                        break;
                    case "corrisk quote follow ups":
                        query = @"
";
                        break;
                    default:
                        query = "";
                        MessageBox.Show("Error! Query not in list.");
                        break;
                }

                SqlCommand getInsureds = new SqlCommand(query, dbConn);

                SqlDataReader returnInsureds = getInsureds.ExecuteReader();
                if (!returnInsureds.HasRows)
                {
                    MessageBox.Show("The query did not return any results");
                    throw new EmptyQueryException();
                }
                else
                {
                    while (returnInsureds.Read())
                    {
                        insuredStorage.Add(new IMSEntry(
                            returnInsureds["Control Number"].ToString(),
                            returnInsureds["Note Due Date"].ToString(),
                            returnInsureds["Effective Date"].ToString(),
                            returnInsureds["Created Date"].ToString(),
                            returnInsureds["Named Insured"].ToString(),
                            returnInsureds["Broker Company"].ToString(),
                            returnInsureds["Broker Name"].ToString(),
                            returnInsureds["Broker First Name"].ToString(),
                            returnInsureds["Broker Last Name"].ToString(),
                            returnInsureds["Broker Email"].ToString(),
                            returnInsureds["UW First Name"].ToString(),
                            returnInsureds["UW Last Name"].ToString(),
                            returnInsureds["UW Email"].ToString()
                        ));
                    }
                }
                dbConn.Close();
            }
            catch (SqlException sqlExc)
            {
                DateTime theDate = DateTime.Now;
                MessageBox.Show("An error occured! Technical details:\n\n" + sqlExc.GetType().Name + "\n\n" + sqlExc.Message);
                using (StreamWriter sqlExcWriter = new StreamWriter(errorLogPath + "ErrorLog.txt", true))
                {
                    sqlExcWriter.WriteLine(theDate.ToString("MM-dd-yyyy"));
                    sqlExcWriter.WriteLine(sqlExc.GetType().Name);
                    sqlExcWriter.WriteLine(sqlExc.Message + "\n");
                }
            }
            finally
            {
                dbConn.Close();
            }
            return insuredStorage;
        }

        /// <summary>
        /// Removes the duplicate insureds from the list of insureds provided by the SQL query.
        /// This method removes duplicate insureds based on control numbers, as one insured
        /// can have multiple accounts for different lines of business.
        /// </summary>
        /// <param name="ListOfInsureds">A list of type IMSEntry</param>
        /// <returns>Returns the list of insureds with all duplicates deleted</returns>
        public List<IMSEntry> RemoveDuplicateInsureds(List<IMSEntry> ListOfInsureds)
        {
            // TODO: Future: Replace this nested loop with a LINQ query
            for (int i = 0; i < ListOfInsureds.Count; i++)
            {
                for (int j = i + 1; j < ListOfInsureds.Count; j++)
                {
                    if (ListOfInsureds[i].ControlNum.Equals(ListOfInsureds[j].ControlNum))
                    {
                        ListOfInsureds.RemoveAt(j);
                    }
                }
            }
            return ListOfInsureds;
        }

        /// <summary>
        /// Retrieve the SQL DB connection string from the App.config file
        /// </summary>
        /// <remarks>Unprotects to read, reads string and stores it, then protects the connect string in App.config</remarks>
        /// <returns>SQL DB connection string</returns>
        private string GetConnection(string selectedQuery)
        {
            string dbConnection;
            Configuration cfg = ConfigurationManager.OpenExeConfiguration("FollowUpSharp.exe");
            ConnectionStringsSection connectStr = cfg.GetSection("connectionStrings") as ConnectionStringsSection;

            dbConnection = connectStr.ConnectionStrings["Login"].ConnectionString;
            connectStr.SectionInformation.ProtectSection("RsaProtectedConfigurationProvider");
            connectStr.SectionInformation.ForceSave = true;
            return dbConnection;
        }
    }
}