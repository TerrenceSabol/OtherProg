using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Data;

namespace ImportCandidateFromTRC
{
    static class Program
    {
        #region Configuration
        public static Boolean DebugMode = false;

         public static string xDatabaseUsername = "DotNetApp";
         public static string xDatabasePassword = "2dOI64LCazsbQzFtCWpb";
         public static string  DatabaseUsername = "sa";
         public static string  DatabasePassword = "jenzadmin";
        public static string DatabaseHostname = "ncsql1";
        public static string DatabaseCatalog = "TMSEPly";

        /// <summary>
        /// CSV column titles and numbers
        /// </summary>
        public static string[,] ColumnIndex = { 
            //Configure these to match the 0-based column locations
            //This is here in case the format changes again next year, we can simply modify the locations index
            {"ScanDateTime","0"},
            {"AssetNum","1"},
            {"AttendeeID","2"},
            {"Firstname","3"},
            {"MiddleName","4"},
            {"LastName","5"},
            {"Address1","6"},
            {"Address2","7"},
            {"City","8"},
            {"Borough","9"},
            {"County","10"},
            {"State","11"},
            {"Zip","12"},
            {"Country","13"},
            {"Gender","14"},
            {"BirthDate","15"},
            {"SSN","16"},
            {"Ethnicity","17"},
            {"Citizenship","18"},
            {"Religion","19"},
            {"HomePhone","20"},
            {"CellPhone","21"},
            {"TextOK","22"},
            {"Email","23"},
            {"AcademicInterestCD1","24"},
            {"AcademicInterestCD1Act","25"},
            {"AcademicInterest1","26"},
            {"AcademicInterestCD2","27"},
            {"AcademicInterestCD2Act","28"},
            {"AcademicInterest2","29"},
            {"AcademicInterest3","30"},
            {"ExtracurricularActivity","31"},
            {"IntendedMajor","32"},
            {"OtherInterests","33"},
            {"CareerObjective","34"},
            {"SchoolName","35"},
            {"CEEBCode","36"},
            {"SchoolCity","37"},
            {"HSGradStatus","38"},
            {"HSGradYear","39"},
            {"GPA","40"},
            {"ACTEnglish","41"},
            {"ACTMath","42"},
            {"ACTComposite","43"},
            {"SATEnglish","44"},
            {"SATMath","45"},
            {"SATComposite","46"},
            {"SAT_ACTScores","47"},
            {"StartingTerm","48"},
            
        };

        #endregion

        /// <summary>
        /// Gets index of the column based on column title
        /// </summary>
        /// <param name="searchString">Column title</param>
        /// <returns>0-based index of column</returns>
        public static int GetColumnIndexOf(string searchString)
        {
            for (int i = 0; i <= ColumnIndex.GetUpperBound(0); i++)
            {
                if (ColumnIndex[i, 0] == searchString)
                {
                    return int.Parse(ColumnIndex[i, 1]);
                }
            }
            //if we got this far, there are no matches
            return -1;
        }

        /// <summary>
        /// Gets name of column based on index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static string GetColumnTitleOf(int index)
        {
            for (int i = 0; i <= ColumnIndex.GetUpperBound(0); i++)
            {
                if (ColumnIndex[i, 1] == index.ToString())
                {
                    return ColumnIndex[i, 0];
                }
            }
            //if we got this far, there are no matches
            return "Error";
        }

        /// <summary>
        /// Pulls info from CSV file and stores it as list of string arrays
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string[]> parseCSV(string path)
        {
            //
            List<string[]> parsedData = new List<string[]>();

            try
            {
                using (StreamReader readFile = new StreamReader(path))
                {
                    string line;
                    string[] row;
                    string pattern = ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";  //Should be commas that are not encapsulated in quotation marks
                    Regex r = new Regex(pattern);
                    while ((line = readFile.ReadLine()) != null)
                    {
                        row = r.Split(line);
                        parsedData.Add(row);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }


            return parsedData;
        }

        /// <summary>
        /// Checks to see if header titles are what the program is expecting
        /// </summary>
        /// <param name="path"></param>
        /// <returns>True if first line is correct</returns>
        public static Boolean CheckFirstLine(string path)
        {
            string ColumnContent = "";
            int ColumnCount = 0;
            List<string[]> parsedData = parseCSV(path);
            string[] firstRow = parsedData[0];
            foreach (string strColumn in firstRow)
            {
                ColumnContent = RemoveLeadingAndTrailingQuotationMarks(strColumn);
                if (ColumnContent == GetColumnTitleOf(ColumnCount))
                {
                    //great
                }
                else
                {
                    //not so great
                    return false;
                }
                ColumnCount++;
            }

            return true;
        }

        /// <summary>
        /// With a name like that, does it really need a summary?
        /// </summary>
        /// <param name="theString"></param>
        /// <returns></returns>
        public static string RemoveLeadingAndTrailingQuotationMarks(string theString)
        {
            if (theString.StartsWith("\""))
            {
                theString = theString.Remove(0, 1);
            }

            if (theString.EndsWith("\""))
            {
                theString = theString.Remove(theString.Length - 1);
            }

            return theString;
        }

        public static string RemoveAllQuotes(string theString)
        {
            string theNewString;
            theNewString = theString.Replace("'", "");
            theNewString = theNewString.Replace("\"", "");
            return theNewString;
        }

        public static string RemoveNonNumeric(string theString)
        {
            return Regex.Replace(theString, "[^0-9]", "");
        }

        public static string RemoveNonAlphanumeric(string theString)
        {
            return Regex.Replace(theString, @"[^\w\s]", string.Empty);
        }

        public static void ExecuteSQLStatement(string strSQL)
        {
            SqlConnection cn = new SqlConnection("Data Source=" + DatabaseHostname + ";Initial Catalog=" + DatabaseCatalog + ";User Id=" + DatabaseUsername + ";Password=" + DatabasePassword + ";");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = strSQL;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = cn;
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();
        }

        public static string GetNextSeqNumber()
        {
            SqlConnection cn = new SqlConnection("Data Source=" + DatabaseHostname + ";Initial Catalog=" + DatabaseCatalog + ";User Id=" + DatabaseUsername + ";Password=" + DatabasePassword + ";");
            SqlCommand cmd = new SqlCommand("[adm].[GetNextSeqNumber]", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlParameter parm = new SqlParameter("@REF_ID", SqlDbType.VarChar);
            parm.Size = 300;
            parm.Value = "NAME_MASTER";
            parm.Direction = ParameterDirection.Input;
            cmd.Parameters.Add(parm);
            SqlParameter parm2 = new SqlParameter("@NextSeqNum", SqlDbType.Int);
            //parm2.Size = 50;
            parm2.Direction = ParameterDirection.Output; // This is important!
            cmd.Parameters.Add(parm2);
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();

            return cmd.Parameters["@NextSeqNum"].Value.ToString();
        }

        public static string GetSingleResult(string SQLquery)
        {
            string theResult = "";

            SqlConnection sqlConnection1 = new SqlConnection("Data Source=" + DatabaseHostname + ";Initial Catalog=" + DatabaseCatalog + ";User Id=" + DatabaseUsername + ";Password=" + DatabasePassword + ";");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
            cmd.CommandText = SQLquery;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            reader = cmd.ExecuteReader();
            while (reader.HasRows)
            {
                while (reader.Read())
                {
                    theResult = reader.GetValue(0).ToString();
                }
                reader.NextResult();
            }
            sqlConnection1.Close();
            return theResult;
        }

        public static void WriteToFileLog(string strLog)
        {
            string strFileName = System.AppDomain.CurrentDomain.FriendlyName + ".log.txt";
            using (StreamWriter w = File.AppendText(strFileName))
            {
                w.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + strLog);
                w.WriteLine("----------");
            }
        }


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

    }
}
