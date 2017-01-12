using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace mailMergeBeta
{
    /// <summary>
    /// Query class that connects to database and fetches QFU info</summary>
    /// <remarks>
    /// This class allows for connection to the specified database and
    /// provides a number of methods for fetching the necessary information
    /// for Quote Follow Ups as well as simply printing them to console if need
    /// be.</remarks>
    class Query
    {
        /// <summary>
        /// Global class variable for interacting with the connection
        /// to the SQL database.</summary>
        private SqlConnection dbConn;

        /// <summary>
        /// The class constructor. Currently takes no arguments,
        /// might change in the future.</summary>
        public Query()
        {
            dbConn = new SqlConnection(
                "server = WINDOWS-KUGQ4HC\\TESTDATABASE;" +
                "Trusted_Connection = yes;" +
                "database = datapile;" +
                "connection timeout = 30"
                );
        }

        /// <summary>
        /// Fetches the list of active control numbers based on the QFU query.
        /// </summary>
        /// <returns>Returns a string list of control numbers.</returns>
        public List<string> fetchCtrlNum()
        {
            dbConn.Open();
            List<string> ctrlStorage = new List<string>();
            string numQuery = "select ctrlNum from followup";
            SqlCommand getCtrlNum = new SqlCommand(numQuery, dbConn);

            SqlDataReader returnCtrlNum = getCtrlNum.ExecuteReader();
            while (returnCtrlNum.Read())
            {
                ctrlStorage.Add(returnCtrlNum["ctrlNum"].ToString());
            }
            dbConn.Close();
            return ctrlStorage;
        }

        /// <summary>
        /// Fetches the list of active names associated with said control numbers
        /// based on the QFU query.</summary>
        /// <returns>Returns a string list of broker names.</returns>
        public List<string> fetchNames()
        {
            dbConn.Open();
            List<string> nameStorage = new List<string>();
            string nameQuery = "select firstName from followup";
            SqlCommand getName = new SqlCommand(nameQuery, dbConn);

            SqlDataReader returnName = getName.ExecuteReader();
            while (returnName.Read())
            {
                nameStorage.Add(returnName["firstName"].ToString());
            }
            dbConn.Close();
            return nameStorage;
        }

        /// <summary>
        /// Fetches the list of active emails associated with said control numbers
        /// based on the QFU query.</summary>
        /// <returns>Returns a string list of broker emails.</returns>
        public List<string> fetchEmails()
        {
            dbConn.Open();
            List<string> emailStorage = new List<string>();
            string emailQuery = "select email from followup";
            SqlCommand getEmail = new SqlCommand(emailQuery, dbConn);

            SqlDataReader returnEmail = getEmail.ExecuteReader();
            while (returnEmail.Read())
            {
                emailStorage.Add(returnEmail["email"].ToString());
            }
            dbConn.Close();
            return emailStorage;
        }

        /// <summary>
        /// Prints the active control numbers to stdout.</summary>
        public void getCtrlNums()
        {
            dbConn.Open();
            string numQuery = "select ctrlNum from followup";
            SqlCommand getCtrlNum = new SqlCommand(numQuery, dbConn);

            SqlDataReader returnCtrlNum = getCtrlNum.ExecuteReader();
            while (returnCtrlNum.Read())
            {
                Console.WriteLine(returnCtrlNum["ctrlNum"].ToString());
            }
            dbConn.Close();
        }

        /// <summary>
        /// Prints the active broker names to stdout.</summary>
        public void getNames()
        {
            dbConn.Open();
            string nameQuery = "select firstName from followup";
            SqlCommand getName = new SqlCommand(nameQuery, dbConn);

            SqlDataReader returnName = getName.ExecuteReader();
            while (returnName.Read())
            {
                Console.WriteLine(returnName["firstName"].ToString());
            }
            dbConn.Close();
        }

        /// <summary>
        /// Prints the active emails to stdout.</summary>
        public void getEmails()
        {
            dbConn.Open();
            string emailQuery = "select email from followup";
            SqlCommand getEmail = new SqlCommand(emailQuery, dbConn);

            SqlDataReader returnEmail = getEmail.ExecuteReader();
            while (returnEmail.Read())
            {
                Console.WriteLine(returnEmail["email"].ToString());
            }
            dbConn.Close();
        }

        // TODO: Get insured names from SQL query
        // TODO: Get effective dates from SQL query
    }
}
