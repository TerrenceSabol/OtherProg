using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportCandidateFromTRC
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnImport.Enabled = false;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //
            
            //

            openFileDialog1.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog1.Multiselect = false;
            openFileDialog1.ShowDialog();
            txtFileLoc.Text = openFileDialog1.FileName;
            btnImport.Enabled = true;
            txtCurTerm.Focus();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try//to validate term and year
            {
                int CurTerm = int.Parse(txtCurTerm.Text);
                int CurYear = int.Parse(txtCurYear.Text);

                if ((CurTerm < 10) || (CurTerm > 90))
                {
                    MessageBox.Show("The value entered for Current Term is not valid!");
                    return;
                }
                if ((CurYear < 2011) || (CurYear > 2025))
                {
                    MessageBox.Show("The value entered for Current Year is not valid!");
                    return;
                }

            }
            catch// if simply parsing them failed
            {
                MessageBox.Show("The values entered for Current Term and Current Year are not valid!");
                return;
            }
            
            
            txtLog.AppendText("--Beginning parse of " + txtFileLoc.Text + "--\n");
            Program.WriteToFileLog("Beginning parse of " + txtFileLoc.Text);
            if (Program.CheckFirstLine(txtFileLoc.Text))
            {
                int numberOfLines;
                txtLog.AppendText("CSV Header Check Passed...\n");
                Program.WriteToFileLog("CSV Header Check Passed");


                //Pull file into list variable
                List<string[]> parsedFile = Program.parseCSV(txtFileLoc.Text);
                numberOfLines = parsedFile.Count - 1;//minus one because the first line is the header

                txtLog.AppendText("Found " + numberOfLines.ToString() + " valid rows in the CSV...\n");
                Program.WriteToFileLog("Found " + numberOfLines.ToString() + " valid rows in the CSV");
                MessageBox.Show("Found " + numberOfLines.ToString() + " entries!");

                for (int i = 1; i <= numberOfLines; i++) //start at one because the first(0) line is headers
                {
                    ParseLine(parsedFile, i);
                }
                txtLog.AppendText("Finished!\n");
                Program.WriteToFileLog("Finished!");
                btnImport.Enabled = false;
            }
            else
            {
                txtLog.AppendText("CSV Header Check Failed!\n");
                Program.WriteToFileLog("CSV Header Check Failed!");
                MessageBox.Show("The CSV is not in the format the program was expecting. Please contact your System Administrator");
            }
        }

        private void ParseLine(List<string[]> parsedFile, int whichLine)
        {
            txtLog.AppendText("Attempting to parse line " + whichLine.ToString() + "...\n");
            Program.WriteToFileLog("Attempting to parse line " + whichLine.ToString() + " of the CSV");
            string[] thisLine = parsedFile[whichLine];
            //MessageBox.Show("Firstname of person number " + whichLine.ToString() + " is " + Program.RemoveLeadingAndTrailingQuotationMarks(thisLine[Program.GetColumnIndexOf("Firstname")])); //Just to prove it works

            //Common stuff
            Program.WriteToFileLog("Attempting to get next available ID number");
            string sqlIDNUM = "";
            try
            {
                sqlIDNUM = Program.GetNextSeqNumber();
            }
            catch (Exception ex)
            {
                Program.WriteToFileLog("Attempt Failed!");
                Program.WriteToFileLog(ex.ToString());
                MessageBox.Show("An error has occured. Please see the application log. The program will now terminate.");
                System.Diagnostics.Process.GetCurrentProcess().Kill(); //TODO: Terrible way of killing the program
            }
            string sqlUSERNAME = "PurchasedList02";
            string sqlJOBNAME = "ImportCandidateFromTRC";
            string sqlJOBTIME = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            //Name_Master stuff
            string sqlLASTNAME = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("LastName")]);
            string sqlFIRSTNAME = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Firstname")]);
            string sqlMIDDLENAME = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("MiddleName")]);
            string sqlMOBILEPHONE = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("CellPhone")]);
            string sqlEMAILADDRESS = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Email")]);

            //Address_Master stuff
            string sqlADDRCDE = "*LPH";
            string sqlADDRLINE1 = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Address1")]);
            string sqlCITY = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("City")]);
            string sqlSTATE = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("State")]);
            string sqlZIP = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Zip")]);
            string sqlCOUNTY = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("County")]);

            //Biograph_Master stuff
            string sqlGENDER = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Gender")]);
            string sqlSSN = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("SSN")]);
            string sqlBIRTHDTE = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("BirthDate")]);
            string sqlETHNICGROUP = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Ethnicity")]);
            string sqlRELIGION = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("Religion")]);

            //Candidate stuff
            string sqlSOURCE5 = "ADPCF";// ???
            string sqlSOURCEDTE5 = sqlJOBTIME;
            string sqlGRADYRLASTORG = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("HSGradYear")]);
            string sqlHIGHSCHOOL = Program.RemoveAllQuotes(thisLine[Program.GetColumnIndexOf("SchoolName")]);
            string sqlCOUNSELORINITIALS = "";//can be NULL
            string sqlRESIDENTOFSTATE = sqlSTATE;
            string sqlCURCANDTYPE = "F";// F for freshman
            string sqlCURDEPT = "";// ??? NULL
            string sqlCURDIV = "UG";
            string sqlCURLOC = "MAIN";
            string sqlCURPROG = "";// ??? NULL
            string sqlCURSTAGE = "02";// TODO: 02 for HS fresh, 03 for HS soph, 04 for HS jun, 05 for HS sen
            string sqlCURTRM = "";// ??? NULL
            string sqlCURYR = "";// ??? NULL

            //Candidacy stuff
            string sqlCANDIDACYTYPE = "F";// F for freshman
            string sqlDEPTCDE = "";// NULL
            string sqlDIVCDE = "UG";
            string sqlLOCACDE = "MAIN";// ???
            string sqlPROGCDE = "UND";// Undetermined
            string sqlSTAGE = sqlCURSTAGE;
            string sqlTRMCDE = txtCurTerm.Text;
            string sqlYRCDE = txtCurYear.Text;
            string sqlHISTSTAGEDTE = sqlJOBTIME;
            string sqlCURCANDIDACY = "Y";// no idea




            InsertIntoNameMaster(sqlIDNUM, sqlLASTNAME, sqlFIRSTNAME, sqlMIDDLENAME, sqlMOBILEPHONE, sqlEMAILADDRESS, sqlUSERNAME, sqlJOBNAME, sqlJOBTIME);
            InsertIntoAddressMaster(sqlIDNUM, sqlADDRCDE, sqlADDRLINE1, sqlCITY, sqlSTATE, sqlZIP, sqlCOUNTY, sqlUSERNAME, sqlJOBNAME, sqlJOBTIME);
            InsertIntoBiographMaster(sqlIDNUM, sqlGENDER, sqlSSN, sqlBIRTHDTE, sqlETHNICGROUP, sqlRELIGION, sqlUSERNAME, sqlJOBNAME, sqlJOBTIME);
            InsertIntoCandidate(sqlIDNUM, sqlSOURCE5, sqlSOURCEDTE5, sqlGRADYRLASTORG, sqlHIGHSCHOOL, sqlCOUNSELORINITIALS, sqlRESIDENTOFSTATE, sqlCURCANDTYPE, sqlCURDEPT, sqlCURDIV, sqlCURLOC, sqlCURPROG, sqlCURSTAGE, sqlCURTRM, sqlCURYR, sqlUSERNAME, sqlJOBNAME, sqlJOBTIME);
            InsertIntoCandidacy(sqlIDNUM, sqlCANDIDACYTYPE, sqlDEPTCDE, sqlDIVCDE, sqlLOCACDE, sqlPROGCDE, sqlSTAGE, sqlTRMCDE, sqlYRCDE, sqlHISTSTAGEDTE, sqlCURCANDIDACY, sqlUSERNAME, sqlJOBNAME, sqlJOBTIME);


        }

        private void InsertIntoNameMaster(string strIDNUM, string strLASTNAME, string strFIRSTNAME, string strMIDDLENAME, string strMOBILEPHONE, string strEMAILADDRESS, string strUSERNAME, string strJOBNAME, string strJOBTIME)
        {
            //NOTE! This function is expecting the variables to be pre-scrubbed!
            string theSQL = "";

            //format strMOBILEPHONE
            strMOBILEPHONE = Program.RemoveNonNumeric(strMOBILEPHONE);
            if (strMOBILEPHONE == "")
            {
                strMOBILEPHONE = "NULL";
            }

            theSQL += "INSERT INTO [dbo].[NAME_MASTER] ([ID_NUM], [LAST_NAME], [FIRST_NAME], [MIDDLE_NAME], [MOBILE_PHONE], [EMAIL_ADDRESS], [USER_NAME], [JOB_NAME], [JOB_TIME]) VALUES (";
            theSQL += "'" + strIDNUM + "', ";
            theSQL += "'" + strLASTNAME + "', ";
            theSQL += "'" + strFIRSTNAME + "', ";
            theSQL += "'" + strMIDDLENAME + "', ";
            theSQL += "" + strMOBILEPHONE + ", ";
            theSQL += "'" + strEMAILADDRESS + "', ";
            theSQL += "'" + strUSERNAME + "', ";
            theSQL += "'" + strJOBNAME + "', ";
            theSQL += "'" + strJOBTIME + "');";


            if (Program.DebugMode)
            {
                txtLog.AppendText(theSQL + "\n");
                txtLog.AppendText("NAME_MASTER record write simulated successfully...\n");
                Program.WriteToFileLog("NAME_MASTER record write simulated successfully");
            }
            else
            {
                ExecuteStatement(theSQL);
                txtLog.AppendText("NAME_MASTER record written successfully...\n");
                Program.WriteToFileLog("NAME_MASTER record written successfully");
            }
        }

        private void InsertIntoAddressMaster(string strIDNUM, string strADDRCDE, string strADDRLINE1, string strCITY, string strSTATE, string strZIP, string strCOUNTY, string strUSERNAME, string strJOBNAME, string strJOBTIME)
        {
            //NOTE! This function is expecting the variables to be pre-scrubbed!
            string theSQL = "";

            theSQL += "INSERT INTO [dbo].[ADDRESS_MASTER] ([ID_NUM],[ADDR_CDE],[ADDR_LINE_1],[CITY],[STATE],[ZIP],[COUNTY],[USER_NAME],[JOB_NAME],[JOB_TIME]) VALUES (";
            theSQL += "'" + strIDNUM + "', ";
            theSQL += "'" + strADDRCDE + "', ";
            theSQL += "'" + strADDRLINE1 + "', ";
            theSQL += "'" + strCITY + "', ";
            theSQL += "'" + strSTATE + "', ";
            theSQL += "'" + strZIP + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strCOUNTY + "', "; //Always set county to null because county is a code and I don't feel like looking those up
            theSQL += "'" + strUSERNAME + "', ";
            theSQL += "'" + strJOBNAME + "', ";
            theSQL += "'" + strJOBTIME + "');";


            if (Program.DebugMode)
            {
                txtLog.AppendText(theSQL + "\n");
                txtLog.AppendText("ADDRESS_MASTER record write simulated successfully...\n");
                Program.WriteToFileLog("ADDRESS_MASTER record write simulated successfully");
            }
            else
            {
                ExecuteStatement(theSQL);
                txtLog.AppendText("ADDRESS_MASTER record written successfully...\n");
                Program.WriteToFileLog("ADDRESS_MASTER record written successfully");
            }
        }

        private void InsertIntoBiographMaster(string strIDNUM, string strGENDER, string strSSN, string strBIRTHDTE, string strETHNICGRP, string strRELIGION, string strUSERNAME, string strJOBNAME, string strJOBTIME)
        {
            //NOTE! This function is expecting the variables to be pre-scrubbed!
            string theSQL = "";

            //format strSSN
            strSSN = Program.RemoveNonNumeric(strSSN);
            if (strSSN == "")
            {
                strSSN = "NULL";
            }
            //format strBIRTHDTE
            try
            {
                DateTime dtmBirth = DateTime.ParseExact(strBIRTHDTE, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                strBIRTHDTE = "'" + dtmBirth.ToString("yyyy-MM-dd HH:mm:ss") + "'";
            }
            catch
            {
                strBIRTHDTE = "NULL";
            }

            theSQL += "INSERT INTO [dbo].[BIOGRAPH_MASTER] ([ID_NUM],[GENDER],[BIRTH_DTE],[ETHNIC_GROUP],[RELIGION],[USER_NAME],[JOB_NAME],[JOB_TIME]) VALUES (";
            theSQL += "'" + strIDNUM + "', ";
            theSQL += "'" + strGENDER + "', ";
            
            theSQL += "" + strBIRTHDTE + ", ";
            theSQL += "NULL, "; //theSQL += "'" + strETHNICGRP + "', ";//Always set this to null because it is a code and I don't feel like looking those up
            theSQL += "NULL, "; //theSQL += "'" + strRELIGION + "', ";//Always set this to null because it is a code and I don't feel like looking those up
            theSQL += "'" + strUSERNAME + "', ";
            theSQL += "'" + strJOBNAME + "', ";
            theSQL += "'" + strJOBTIME + "');";


            if (Program.DebugMode)
            {
                txtLog.AppendText(theSQL + "\n");
                txtLog.AppendText("BIOGRAPH_MASTER record write simulated successfully...\n");
                Program.WriteToFileLog("BIOGRAPH_MASTER record write simulated successfully");
            }
            else
            {
                ExecuteStatement(theSQL);
                txtLog.AppendText("BIOGRAPH_MASTER record written successfully...\n");
                Program.WriteToFileLog("BIOGRAPH_MASTER record written successfully");
            }
        }

        private void InsertIntoCandidate(string strIDNUM, string strSOURCE5, string strSOURCEDTE5, string strGRADYRLASTORG, string strHIGHSCHOOL, string strCOUNSELORINITIALS, string strRESIDENTOFSTATE, string strCURCANDTYPE, string strCURDEPT, string strCURDIV, string strCURLOC, string strCURPROG, string strCURSTAGE, string strCURTRM, string strCURYR, string strUSERNAME, string strJOBNAME, string strJOBTIME)
        {
            //NOTE! This function is expecting the variables to be pre-scrubbed!
            string theSQL = "";

            //Format strHIGHSCHOOL
            strHIGHSCHOOL = Program.RemoveNonAlphanumeric(strHIGHSCHOOL);
            strHIGHSCHOOL = GetHSCode(strHIGHSCHOOL);
            if (strHIGHSCHOOL == "")
            {
                strHIGHSCHOOL = "NULL";
            }
            else
            {
                strHIGHSCHOOL = "'" + strHIGHSCHOOL + "'";
            }

            theSQL += "INSERT INTO [dbo].[CANDIDATE] ([ID_NUM],[SOURCE_5],[SOURCE_DTE_5], [GRAD_YR_LAST_ORG],[HIGH_SCHOOL], [COUNSELOR_INITIALS], [RESIDENT_OF_STATE],[CUR_CAND_TYPE], [CUR_DEPT], [CUR_DIV], [CUR_LOC],[CUR_PROG], [CUR_STAGE], [CUR_TRM], [CUR_YR],[USER_NAME],[JOB_NAME],[JOB_TIME] ) VALUES (";
            theSQL += "'" + strIDNUM + "', ";
            theSQL += "'" + strSOURCE5 + "', ";
            theSQL += "'" + strSOURCEDTE5 + "', ";
            theSQL += "'" + strGRADYRLASTORG + "', ";
            theSQL += "" + strHIGHSCHOOL + ", ";
            theSQL += "NULL, "; //theSQL += "'" + strCOUNSELORINITIALS + "', ";
            theSQL += "'" + strRESIDENTOFSTATE + "', ";
            theSQL += "'" + strCURCANDTYPE + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strCURDEPT + "', ";
            theSQL += "'" + strCURDIV + "', ";
            theSQL += "'" + strCURLOC + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strCURPROG + "', ";
            theSQL += "'" + strCURSTAGE + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strCURTRM + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strCURYR + "', ";
            theSQL += "'" + strUSERNAME + "', ";
            theSQL += "'" + strJOBNAME + "', ";
            theSQL += "'" + strJOBTIME + "');";


            if (Program.DebugMode)
            {
                txtLog.AppendText(theSQL + "\n");
                txtLog.AppendText("CANDIDATE record write simulated successfully...\n");
                Program.WriteToFileLog("CANDIDATE record write simulated successfully");
            }
            else
            {
                ExecuteStatement(theSQL);
                txtLog.AppendText("CANDIDATE record written successfully...\n");
                Program.WriteToFileLog("CANDIDATE record written successfully");
            }
        }

        private void InsertIntoCandidacy(string strIDNUM, string strCANDIDACYTYPE, string strDEPTCDE, string strDIVCDE, string strLOCACDE, string strPROGCDE, string strSTAGE, string strTRMCDE, string strYRCDE, string strHISTSTAGE, string strCURCANDIDACY, string strUsername, string strJOBNAME, string strJOBTIME)
        {
            //NOTE! This function is expecting the variables to be pre-scrubbed!
            string theSQL = "";

            theSQL += "INSERT INTO [dbo].[CANDIDACY] ([ID_NUM],[CANDIDACY_TYPE],[DEPT_CDE],[DIV_CDE],[LOCA_CDE],[PROG_CDE], [STAGE],[TRM_CDE],[YR_CDE],[HIST_STAGE_DTE],[CUR_CANDIDACY], [USER_NAME],[JOB_NAME],[JOB_TIME]) VALUES (";
            theSQL += "'" + strIDNUM + "', ";
            theSQL += "'" + strCANDIDACYTYPE + "', ";
            theSQL += "NULL, "; //theSQL += "'" + strDEPTCDE + "', ";
            theSQL += "'" + strDIVCDE + "', ";
            theSQL += "'" + strLOCACDE + "', ";
            theSQL += "'" + strPROGCDE + "', ";
            theSQL += "'" + strSTAGE + "', ";
            theSQL += "'" + strTRMCDE + "', ";
            theSQL += "'" + strYRCDE + "', ";
            theSQL += "'" + strHISTSTAGE + "', ";
            theSQL += "'" + strCURCANDIDACY + "', ";
            theSQL += "'" + strUsername + "', ";
            theSQL += "'" + strJOBNAME + "', ";
            theSQL += "'" + strJOBTIME + "');";


            if (Program.DebugMode)
            {
                txtLog.AppendText(theSQL + "\n");
                txtLog.AppendText("CANDIDACY record write simulated successfully...\n");
                Program.WriteToFileLog("CANDIDACY record write simulated successfully");
            }
            else
            {
                ExecuteStatement(theSQL);
                txtLog.AppendText("CANDIDACY record written successfully...\n");
                Program.WriteToFileLog("CANDIDACY record written successfully");
            }
        }

        private void ExecuteStatement(string theSQL)
        {
            Program.WriteToFileLog("Attempting to execute the following SQL:");
            Program.WriteToFileLog(theSQL);
            
            try
            {
                Program.ExecuteSQLStatement(theSQL);
                Program.WriteToFileLog("Success");
            }
            catch (Exception ex)
            {
                Program.WriteToFileLog("Failed!");
                Program.WriteToFileLog(ex.ToString());
                MessageBox.Show("An error has occured. Please see the application log. The program will now terminate.");
                System.Diagnostics.Process.GetCurrentProcess().Kill(); //Terrible way of killing the program
            }
        }

        private string GetHSCode(string searchString)
        {
            return Program.GetSingleResult("SELECT n.[ID_NUM] FROM [dbo].[NAME_MASTER] AS n JOIN [dbo].[ORG_MASTER] AS o ON o.ID_NUM = n.ID_NUM WHERE n.LAST_NAME LIKE '%" + searchString + "%' AND o.ORG_TYPE = 'HS'");
        }

    }
}
