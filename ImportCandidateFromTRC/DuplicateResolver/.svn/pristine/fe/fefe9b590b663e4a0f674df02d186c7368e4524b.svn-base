using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DuplicateResolver
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Hide();
            progressBar1.Location = btnResolve.Location;
            progressBar1.Size = btnResolve.Size;

            Program.WriteLog("Program opened by " + Environment.UserName);
        }

        private void btnResolve_Click(object sender, System.EventArgs e)
        {
            int HowManyHolding = NumberOfHoldingRecords();
            if (HowManyHolding > 0)
            {
                MessageBox.Show("Found " + HowManyHolding + " candidates in the holding records. Click OK to resolve duplicates.");

                progressBar1.Show();
                btnResolve.Hide();
                ResolveDuplicates(HowManyHolding);
                btnResolve.Show();
                progressBar1.Hide();
            }
            else
            {
                MessageBox.Show("There are no candidates in the holding records!");
            }
        }

        /// <summary>
        /// Cycle through Holding records X times to resolve them
        /// </summary>
        /// <param name="HowManyToResolve">X</param>
        private void ResolveDuplicates(int HowManyToResolve)
        {
            progressBar1.Maximum = HowManyToResolve * 7; //7 steps per person
            progressBar1.Value = 0;

            //create query to get next person from holding tables
            SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            string sqlQuery = "SELECT TOP 1 [ID_NUM] FROM [dbo].[NC_HOLDING_NAME_MASTER] ORDER BY [ID_NUM] ASC";
            SqlCommand cmd = new SqlCommand(sqlQuery, cn);

            //run it X many times
            for (int i = 1; i <= HowManyToResolve; i++)
            {
                cn.Open();
                ResolveDuplicate(Convert.ToInt32(cmd.ExecuteScalar()));
                cn.Close();
            }
        }

        private void ResolveDuplicate(int HoldingIDNumber)
        {
            Program.WriteLog("Attempting to resolve candidate with holding ID of " + HoldingIDNumber.ToString());
            int ExistingJenzID = IsThisADuplicate(HoldingIDNumber);
            if (ExistingJenzID == 0)
            {
                //Not a duplicate. Simply write to tables.
                Program.WriteLog("Copying person with ID of " + HoldingIDNumber + " to Jenzabar tables.");
                MoveAllRecords(HoldingIDNumber);
                
            }
            else
            {
                //is a duplicate. Resolve columns
                Program.WriteLog("Resolving columns with holdingID = " + HoldingIDNumber + " and JenzID = " + ExistingJenzID + ".");
                ResolveAllColumns(HoldingIDNumber, ExistingJenzID);
                Program.WriteLog("Columns have been resolved.");
            }

            //Finally, delete them from holding tables.
            Program.WriteLog("Deleting holding ID of " + HoldingIDNumber + ".");
            DeleteHoldingRecords(HoldingIDNumber);
        }

        /// <summary>
        /// Resolve columns of duplicate IDs
        /// </summary>
        /// <param name="HoldingID"></param>
        /// <param name="ExistingID"></param>
        private void ResolveAllColumns(int HoldingID, int ExistingID)
        {
            CompareTables("NAME_MASTER", ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);

            CompareTables("BIOGRAPH_MASTER", ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);

            CompareTables("CANDIDACY", ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);

            CompareTables("CANDIDATE", ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);

            CompareAddresses(ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);

            //TODO: Compare Attribute trans
            progressBar1.Value++;
            Program.MessyRest(1);

            CompareEthnic(ExistingID, HoldingID);
            progressBar1.Value++;
            Program.MessyRest(1);
        }

        private void CompareTables(string TableName, int ExistingID, int IncomingID)
        {
            string ExistingTable = TableName;
            string IncomingTable = "NC_HOLDING_" + TableName;
            
            SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            Program.WriteLog("Checking [" + ExistingTable + "] vs [" + IncomingTable +  "] tables...");

            string sqlQuery = "SELECT [COLUMN_NAME] FROM [INFORMATION_SCHEMA].[COLUMNS] ";
            sqlQuery += "WHERE [TABLE_NAME] = '" + IncomingTable + "' ";
            sqlQuery += "AND [COLUMN_NAME] NOT LIKE 'ID_NUM' ";
            sqlQuery += "AND [COLUMN_NAME] NOT LIKE 'USER_NAME' ";
            sqlQuery += "AND [COLUMN_NAME] NOT LIKE 'JOB_NAME' ";
            sqlQuery += "AND [COLUMN_NAME] NOT LIKE 'JOB_TIME' ";
            sqlQuery += "AND [COLUMN_NAME] NOT LIKE 'PROG_CDE' ";
            sqlQuery += "ORDER BY [ORDINAL_POSITION]";
            

            cmd.CommandText = sqlQuery;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = cn;
            cn.Open();
            reader = cmd.ExecuteReader();

            //pull data from reader
            while (reader.HasRows)
            {
                while (reader.Read()) //cycle through columns
                {
                    string ColumnName = reader.GetValue(0).ToString();
                    Program.WriteLog("Checking both [" + ColumnName + "] columns for data...");
                    string ExistingColumnValue = GetResult(ExistingTable, ExistingID, ColumnName);
                    string IncomingColumnValue = GetResult(IncomingTable, IncomingID, ColumnName);
                    if (ExistingColumnValue == IncomingColumnValue)
                    {
                        //Then both entries match, there is no updating to be done
                        Program.WriteLog("Columns match.");
                    }
                    else
                    {
                        //They don't match 
                        if (ColumnName.Contains("DTE"))
                        {
                            Program.WriteLog("Column is a date column. Updating...");
                            ExecuteNonQuery("UPDATE [" + ExistingTable + "] SET [" + ColumnName +
                                "] = (SELECT [" + ColumnName + "] FROM [" + IncomingTable + "] WHERE [ID_NUM] = '" +
                                IncomingID + "') WHERE [ID_NUM] = '" + ExistingID + "'");
                        }
                        else if (ExistingColumnValue == "")
                        {
                            Program.WriteLog("Incoming has data, existing does not. Writing new data to Jenz...");
                            
                            ExecuteNonQuery("UPDATE [" + ExistingTable + "] SET [" + ColumnName +
                                "] = (SELECT [" + ColumnName + "] FROM [" + IncomingTable + "] WHERE [ID_NUM] = '" + 
                                IncomingID + "') WHERE [ID_NUM] = '" + ExistingID + "'");
                            
                        }
                        else if (IncomingColumnValue == "")
                        {
                            Program.WriteLog("Incoming is blank, Jenz has existing data. Skipping column.");
                            //don't do anything. Incoming is blank, but data exists in Jenz. Don't write null over existing data.
                        }
                        else
                        {
                            //neither are blank. 
                            Program.WriteLog("Incoming has data, existing has data. Asking user...");
                            //Ask user what to do.
                            string msg = "Data already exists for the " + ColumnName + " record for ID number " + ExistingID + ".\r\n";
                            msg += "Existing Data: " + ExistingColumnValue + "\r\n";
                            msg += "Incoming Data: " + IncomingColumnValue + "\r\n";
                            msg += "\r\n";
                            msg += "Do you wish to overwrite the existing data?";
                            DialogResult dResult = MessageBox.Show(msg, "Modify Existing Data?", MessageBoxButtons.YesNo);
                            if (dResult == DialogResult.Yes)
                            {
                                Program.WriteLog("User wants to overwrite existing data with new data.");
                                //overwrite existing data
                                ExecuteNonQuery("UPDATE [" + ExistingTable + "] SET [" + ColumnName +
                                "] = (SELECT [" + ColumnName + "] FROM [" + IncomingTable + "] WHERE [ID_NUM] = '" +
                                IncomingID + "') WHERE [ID_NUM] = '" + ExistingID + "'");
                            }
                            else
                            {
                                Program.WriteLog("User wants to keep existing data. Skipping column.");
                                //do nothing to existing data.
                            }
                        }
                    }
                }
                reader.NextResult();
            }

            cn.Close();
        }

        private void CompareAddresses(int ExistingID, int HoldingID)
        {
            Program.WriteLog("Checking for addresses in the holding table...");

            SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            string sqlQuery = "SELECT COUNT(*) FROM [NC_HOLDING_ADDRESS_MASTER] where [ID_NUM] = '" + HoldingID + "'";

            SqlCommand cmd = new SqlCommand(sqlQuery, cn);

            cn.Open();
            int HowManyAddresses = Convert.ToInt32(cmd.ExecuteScalar());
            cn.Close();

            Program.WriteLog("Found " + HowManyAddresses + " addresses in the holding table.");
            if (HowManyAddresses > 0)
            {
                //Get incoming addresses
                sqlQuery = "SELECT td.[TABLE_DESC], am.* FROM [NC_HOLDING_ADDRESS_MASTER] AS am " +
                            "LEFT JOIN [TABLE_DETAIL] AS td " +
                            "ON am.[ADDR_CDE] = td.[TABLE_VALUE] " +
                            "WHERE am.[ID_NUM] = '" + HoldingID + "' AND " +
                            "td.[COLUMN_NAME] = 'addr_cde';";
                cmd = new SqlCommand(sqlQuery, cn);
                cn.Open();
                SqlDataReader rdrIncoming = cmd.ExecuteReader();
                //pull data from reader
                while (rdrIncoming.HasRows)
                {
                    while (rdrIncoming.Read()) //cycle through rows of incoming addresses
                    {
                        Program.WriteLog("Found address with code: '" + rdrIncoming.GetValue(2) + " - " + rdrIncoming.GetValue(0) + "'.");
                        Program.WriteLog("Checking to see if an address with this code already exists...");

                        SqlConnection cnB = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
                        string sqlQueryB = "SELECT * FROM [ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';";
                        SqlCommand cmdB = new SqlCommand(sqlQueryB, cnB);
                        cnB.Open();
                        //Should only ever have one row due to ID + ADDR_CDE, so don't have to bother with the while loop BS
                        SqlDataReader rdrExisting = cmdB.ExecuteReader();
                        if (rdrExisting.HasRows && rdrExisting.Read())
                        {
                            Program.WriteLog("Found existing record with same code.");

                            //if (addr_line_1 same) and (city same) and (state same) and (zip same),
                            if ((rdrExisting.GetValue(14).ToString().Trim() == rdrIncoming.GetValue(15).ToString().Trim()) && (rdrExisting.GetValue(17).ToString().Trim() == rdrIncoming.GetValue(18).ToString().Trim()) && (rdrExisting.GetValue(18).ToString().Trim() == rdrIncoming.GetValue(19).ToString().Trim()) && (rdrExisting.GetValue(19).ToString().Trim() == rdrIncoming.GetValue(20).ToString().Trim()))
                            {
                                Program.WriteLog("Addresses are the same.");
                                //don't do anything
                            }
                            else
                            {
                                Program.WriteLog("Addresses are different.");
                                Program.WriteLog("Asking user which to keep...");
                                string msg = "An address with the code " + rdrIncoming.GetValue(2) + " - " + rdrIncoming.GetValue(0) + " already exists.\r\n";
                                msg += "Existing Data: " + rdrExisting.GetValue(14) + ", " + rdrExisting.GetValue(17) + ", " + rdrExisting.GetValue(18) + ", " + rdrExisting.GetValue(19) + "\r\n";
                                msg += "Incoming Data: " + rdrIncoming.GetValue(15) + ", " + rdrIncoming.GetValue(18) + ", " + rdrIncoming.GetValue(19) + ", " + rdrIncoming.GetValue(20) + "\r\n";
                                msg += "\r\n";
                                msg += "Do you wish to overwrite the existing data?";

                                DialogResult dResult = MessageBox.Show(msg, "Modify Existing Data?", MessageBoxButtons.YesNo);
                                if (dResult == DialogResult.Yes)
                                {
                                    Program.WriteLog("User chose to overwrite existing data...");
                                    //delete existing row from address_master
                                    ExecuteNonQuery("DELETE FROM [ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                                    //set incoming ID to match existing
                                    ExecuteNonQuery("UPDATE [NC_HOLDING_ADDRESS_MASTER] SET [ID_NUM] = '" + ExistingID + "' WHERE [ID_NUM] = '" + HoldingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                                    //copy that row to existing table
                                    ExecuteNonQuery("INSERT INTO [ADDRESS_MASTER] SELECT * FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                                    //delete address from holding
                                    ExecuteNonQuery("DELETE FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                                }
                                else
                                {
                                    Program.WriteLog("User chose to keep existing data.");
                                    //don't do anything
                                }
                            }
                        }
                        else
                        {
                            Program.WriteLog("No existing address with this code found. Writing address to Jenzabar DB...");
                            //set incoming ID to match existing
                            ExecuteNonQuery("UPDATE [NC_HOLDING_ADDRESS_MASTER] SET [ID_NUM] = '" + ExistingID + "' WHERE [ID_NUM] = '" + HoldingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                            //copy that row to existing table
                            ExecuteNonQuery("INSERT INTO [ADDRESS_MASTER] SELECT * FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                            //delete address from holding
                            ExecuteNonQuery("DELETE FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + ExistingID + "' AND [ADDR_CDE] = '" + rdrIncoming.GetValue(2) + "';");
                        }
                        cnB.Close();


                    }
                    rdrIncoming.NextResult();
                }
                cn.Close();
            }
            else
            {
                //aint no addresses. We're done
                Program.WriteLog("No addresses found in the holding table.");
            }
        }

        /// <summary>
        /// Adds IPEDS Ethnic records to existing person
        /// </summary>
        /// <param name="ExistingID"></param>
        /// <param name="HoldingID"></param>
        private void CompareEthnic(int ExistingID, int HoldingID)
        {
            Program.WriteLog("Checking Holding tables for IPEDS Ethnicity records...");

            SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            string sqlQuery = "SELECT COUNT(*) FROM [NC_HOLDING_ETHNIC_RACE_REPORT] where [ID_NUM] = '" + HoldingID + "'";

            SqlCommand cmd = new SqlCommand(sqlQuery, cn);

            cn.Open();
            int HowManyHoldingEthnic = Convert.ToInt32(cmd.ExecuteScalar());
            cn.Close();

            if (HowManyHoldingEthnic < 1)
            {
                Program.WriteLog("No IPEDS Ethnicity records found in the holding tables.");
                //nothing to do here
            }
            else
            {
                Program.WriteLog("Found " + HowManyHoldingEthnic.ToString() + " IPEDS Ethnicity records found in the holding tables.");
                //Get highest existing SEQ_NUM
                sqlQuery = "SELECT TOP 1 [SEQ_NUM] FROM [ETHNIC_RACE_REPORT] where [ID_NUM] = '" + ExistingID + "' ORDER BY [SEQ_NUM] DESC";
                cmd = new SqlCommand(sqlQuery, cn);
                cn.Open();
                int HighestExistingEthnic = Convert.ToInt32(cmd.ExecuteScalar());
                cn.Close();
                Program.WriteLog("Found " + HighestExistingEthnic.ToString() + " existing IPEDS Ethnicity entries.");
                Program.WriteLog("Adding holding IPEDS Ethnicity records to existing records...");
                //set incoming ID to match existing
                ExecuteNonQuery("UPDATE [NC_HOLDING_ETHNIC_RACE_REPORT] SET [ID_NUM] = '" + ExistingID + "' WHERE [ID_NUM] = '" + HoldingID + "';");
                //Increase their SEQ_NUMs by highest existing
                ExecuteNonQuery("UPDATE [NC_HOLDING_ETHNIC_RACE_REPORT] SET [SEQ_NUM] = [SEQ_NUM] + " + HighestExistingEthnic.ToString() + " WHERE [ID_NUM] = '" + ExistingID + "';");
                //copy rows to existing table
                ExecuteNonQuery("INSERT INTO [ETHNIC_RACE_REPORT] SELECT * FROM [NC_HOLDING_ETHNIC_RACE_REPORT] WHERE [ID_NUM] = '" + ExistingID + "';");
                //delete rows from holding
                ExecuteNonQuery("DELETE FROM [NC_HOLDING_ETHNIC_RACE_REPORT] WHERE [ID_NUM] = '" + ExistingID + "';");

            }
        }

        /// <summary>
        /// Moves someone from Holding tables to real tables
        /// </summary>
        /// <param name="HoldingID"></param>
        private void MoveAllRecords(int HoldingID)
        {
            ExecuteNonQuery("INSERT INTO [NAME_MASTER] SELECT * FROM [NC_HOLDING_NAME_MASTER] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [ETHNIC_RACE_REPORT] SELECT * FROM [NC_HOLDING_ETHNIC_RACE_REPORT] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [CANDIDATE] SELECT * FROM [NC_HOLDING_CANDIDATE] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [CANDIDACY] SELECT * FROM [NC_HOLDING_CANDIDACY] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [BIOGRAPH_MASTER] SELECT * FROM [NC_HOLDING_BIOGRAPH_MASTER] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [ATTRIBUTE_TRANS] SELECT * FROM [NC_HOLDING_ATTRIBUTE_TRANS] WHERE [ID_NUM] = '" + HoldingID + "' AND ([ATTRIB_CDE] IS NOT NULL) AND ([ATTRIB_CDE] <> 'OTHER');");
            progressBar1.Value++;
            Program.MessyRest(1);
            ExecuteNonQuery("INSERT INTO [ADDRESS_MASTER] SELECT * FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + HoldingID + "';");
            progressBar1.Value++;
            Program.MessyRest(1);
        }

        /// <summary>
        /// Delete person from holding tables
        /// </summary>
        /// <param name="HoldingIDNumber">ID_NUM from holding tables</param>
        private void DeleteHoldingRecords(int HoldingIDNumber)
        {
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_ATTRIBUTE_TRANS] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_BIOGRAPH_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_CANDIDACY] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_CANDIDATE] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_ETHNIC_RACE_REPORT] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
            ExecuteNonQuery("DELETE FROM [NC_HOLDING_NAME_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "'");
        }

        /// <summary>
        /// Attempts to see if person in Holding table is a duplicate of existing person in Jenz system
        /// </summary>
        /// <param name="HoldingIDNumber"></param>
        /// <returns>0 if not duplicate. Existing Jenz ID if duplicate.</returns>
        private int IsThisADuplicate(int HoldingIDNumber)
        {
            Program.WriteLog("Checking to see if candidate with holding ID of " + HoldingIDNumber + " is a duplicate of an existing Jenzabar entry.");
            int MatchingExistingID;

            string HoldingSSN = GetResult("NC_HOLDING_BIOGRAPH_MASTER", HoldingIDNumber, "SSN");
            string HoldingFirstName = GetResult("NC_HOLDING_NAME_MASTER", HoldingIDNumber, "FIRST_NAME");
            string HoldingLastName = GetResult("NC_HOLDING_NAME_MASTER", HoldingIDNumber, "LAST_NAME");

            if (HoldingSSN == "") //no SSN in holding table. Try other things.
            {
                //TODO: Check name
                return CheckNameAndCity(HoldingIDNumber);
            }
            else 
            {
                //Check if someone exists with this SSN
                SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
                string sqlQuery = "SELECT [ID_NUM] FROM [dbo].[BIOGRAPH_MASTER] WHERE [SSN] = '" + HoldingSSN + "'";
                SqlCommand cmd = new SqlCommand(sqlQuery, cn);
                cn.Open();
                MatchingExistingID = Convert.ToInt32(cmd.ExecuteScalar());
                cn.Close();
                if (MatchingExistingID == 0)
                {
                    //TODO: Check name
                    return CheckNameAndCity(HoldingIDNumber);
                }
                else
                {
                    Program.WriteLog("Found existing person with this SSN.");
                    Program.WriteLog("This person matches existing Jenzabar ID = " + MatchingExistingID + ".");
                    return MatchingExistingID;
                }
            }
            
            //Does not appear to be a duplicate.
            Program.WriteLog("Not a duplicate.");
            return 0;
            
        }

        /// <summary>
        /// Tries to find an existing user with matching first name, last name, city, and state from holding tables
        /// </summary>
        /// <param name="HoldingIDNumber"></param>
        /// <returns>Existing user's ID, if exists. 0 otherwise.</returns>
        private int CheckNameAndCity(int HoldingIDNumber)
        {
            int MatchingExistingID;
            Program.WriteLog("Checking if a person exists with this name and city/state...");
            SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            string sqlQuery = "SELECT nm.[ID_NUM], nm.FIRST_NAME, nm.LAST_NAME, am.[CITY], am.[STATE] FROM [dbo].[NAME_MASTER] as nm ";
            sqlQuery += "LEFT JOIN [dbo].[ADDRESS_MASTER] as am on nm.[ID_NUM] = am.[ID_NUM] ";
            sqlQuery += "WHERE nm.[FIRST_NAME] = (SELECT nh.[FIRST_NAME] FROM [dbo].[NC_HOLDING_NAME_MASTER] as nh WHERE nh.[ID_NUM] = '" + HoldingIDNumber + "') ";
            sqlQuery += "AND nm.[LAST_NAME] = (SELECT nh.[LAST_NAME] FROM [dbo].[NC_HOLDING_NAME_MASTER] as nh WHERE nh.[ID_NUM] = '" + HoldingIDNumber + "') ";
            sqlQuery += "AND am.[ADDR_CDE] = '*LHP' ";
            sqlQuery += "AND am.[CITY] = (SELECT ah.[CITY] FROM [dbo].[NC_HOLDING_ADDRESS_MASTER] as ah WHERE ah.[ID_NUM] = '" + HoldingIDNumber + "' AND ah.[ADDR_CDE] = '*LHP') ";
            sqlQuery += "AND am.[STATE] = (SELECT ah.[STATE] FROM [dbo].[NC_HOLDING_ADDRESS_MASTER] as ah WHERE ah.[ID_NUM] = '" + HoldingIDNumber + "' AND ah.[ADDR_CDE] = '*LHP') ";
            SqlCommand cmd = new SqlCommand(sqlQuery, cn);
            cn.Open();
            MatchingExistingID = Convert.ToInt32(cmd.ExecuteScalar());
            cn.Close();
            if (MatchingExistingID == 0){
                Program.WriteLog("No existing records found with this name and city/state combination.");
                return 0;
            }else{
                Program.WriteLog("Someone already exists with this name and city/state combination. Asking user to confirm...");
                string msg = "Someone already exists in Jenzabar with this First Name, Last Name, City, and State.\r\n";
                msg += "Please confirm they are the same people.\r\n";
                msg += "\r\n";
                msg += "Existing Person: ";
                msg += GetResult("NAME_MASTER", MatchingExistingID, "FIRST_NAME").Trim() + " ";
                msg += GetResult("NAME_MASTER", MatchingExistingID, "LAST_NAME").Trim() + " ";
                msg += GetResult("NAME_MASTER", MatchingExistingID, "SUFFIX").Trim() + "\r\n";
                msg += ExecuteQuery("SELECT [ADDR_LINE_1] FROM [dbo].[ADDRESS_MASTER] WHERE [ID_NUM] = '" + MatchingExistingID + "' AND [ADDR_CDE] = '*LHP'").Trim() + ", ";
                msg += ExecuteQuery("SELECT [CITY] FROM [dbo].[ADDRESS_MASTER] WHERE [ID_NUM] = '" + MatchingExistingID + "' AND [ADDR_CDE] = '*LHP'").Trim() + ", ";
                msg += ExecuteQuery("SELECT [STATE] FROM [dbo].[ADDRESS_MASTER] WHERE [ID_NUM] = '" + MatchingExistingID + "' AND [ADDR_CDE] = '*LHP'").Trim() + "\r\n";
                msg += "\r\n";
                msg += "Incoming Person: ";
                msg += GetResult("NC_HOLDING_NAME_MASTER", HoldingIDNumber, "FIRST_NAME").Trim() + " ";
                msg += GetResult("NC_HOLDING_NAME_MASTER", HoldingIDNumber, "LAST_NAME").Trim() + " ";
                msg += GetResult("NC_HOLDING_NAME_MASTER", HoldingIDNumber, "SUFFIX").Trim() + "\r\n";
                msg += ExecuteQuery("SELECT [ADDR_LINE_1] FROM [dbo].[NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "' AND [ADDR_CDE] = '*LHP'").Trim() + ", ";
                msg += ExecuteQuery("SELECT [CITY] FROM [dbo].[NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "' AND [ADDR_CDE] = '*LHP'").Trim() + ", ";
                msg += ExecuteQuery("SELECT [STATE] FROM [dbo].[NC_HOLDING_ADDRESS_MASTER] WHERE [ID_NUM] = '" + HoldingIDNumber + "' AND [ADDR_CDE] = '*LHP'").Trim() + "\r\n";
                msg += "\r\n";
                msg += "Are these the same people?";

                DialogResult dResult = MessageBox.Show(msg, "Suspected Duplicate - Please Confirm", MessageBoxButtons.YesNo);
                if (dResult == DialogResult.Yes)
                {
                    Program.WriteLog("User says these are the same people.");
                    return MatchingExistingID;
                }
                else
                {
                    Program.WriteLog("User says these are not the same people.");
                    return 0;
                }

            }
        }
        

        /// <summary>
        /// Queries the database for number of people in the Holding system
        /// </summary>
        /// <returns>Number of people in the Holding system</returns>
        private int NumberOfHoldingRecords()
        {
            try
            {
                int theAnswer;
                SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
                string sqlQuery = "SELECT COUNT(*) FROM [dbo].[NC_HOLDING_NAME_MASTER]";

                SqlCommand cmd = new SqlCommand(sqlQuery, cn);

                cn.Open();
                theAnswer = Convert.ToInt32(cmd.ExecuteScalar());
                cn.Close();

                return theAnswer;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Execute SQL command
        /// </summary>
        /// <param name="strSQL"></param>
        private void ExecuteNonQuery(string strSQL)
        {
            Program.WriteLog("Executing the following SQL Statement.\r\n" + strSQL);
            SqlConnection conn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
            SqlCommand commd = new SqlCommand();
            commd.CommandText = strSQL;
            commd.CommandType = CommandType.Text;
            commd.Connection = conn;
            conn.Open();
            try
            {
                commd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                Program.WriteLog("Execute Query Failed.\r\n" + e.Message);
            }
            conn.Close();
        }

        private string ExecuteQuery(string strSQL)
        {
            try
            {
                string theAnswer;
                SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
                SqlCommand cmd = new SqlCommand(strSQL, cn);
                cn.Open();
                theAnswer = cmd.ExecuteScalar().ToString();
                cn.Close();
                return theAnswer;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// SELECT strColumn FROM strTable WHERE ID_NUM = IDNum
        /// </summary>
        /// <param name="strTable">Name of SQL Table</param>
        /// <param name="IDNum">WHERE ID_NUM = this</param>
        /// <param name="strColumn">Name of column in SQL Table</param>
        /// <returns>First result of SELECT strColumn FROM strTable WHERE ID_NUM = IDNum </returns>
        private string GetResult(string strTable, int IDNum, string strColumn)
        {
            try
            {
                string  theAnswer;
                SqlConnection cn = new SqlConnection("Data Source=" + Program.DatabaseHostname + ";Initial Catalog=" + Program.DatabaseCatalog + ";User Id=" + Program.DatabaseUsername + ";Password=" + Program.DatabasePassword + ";");
                string sqlQuery = "SELECT [" + strColumn + "] FROM [dbo].[" + strTable + "] WHERE [ID_NUM] = '" + IDNum + "'";

                SqlCommand cmd = new SqlCommand(sqlQuery, cn);

                cn.Open();
                theAnswer = cmd.ExecuteScalar().ToString();
                cn.Close();

                return theAnswer;
            }
            catch
            {
                return "";
            }
        }
    }
}
