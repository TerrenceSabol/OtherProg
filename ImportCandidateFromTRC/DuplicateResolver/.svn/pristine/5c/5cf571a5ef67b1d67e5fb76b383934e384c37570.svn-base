using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DuplicateResolver
{
    static class Program
    {
        public static string DatabaseUsername = "DotNetApp";
        public static string DatabasePassword = "2dOI64LCazsbQzFtCWpb";
        public static string DatabaseHostname = "ncsql1";
        public static string DatabaseCatalog = "TmsEPrd";
        public static string JobName = "DupResolver";



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


        /// <summary>
        /// Give paint events a chance to happen
        /// </summary>
        /// <param name="ms"></param>
        public static void MessyRest(int ms)
        {
            for (int i = 1; i <= ms; i = i + 2) //double increment
            {                                   //...
                Application.DoEvents();         //because this is probably gonna take at least 1ms
                System.Threading.Thread.Sleep(1);
            }
        }

        /// <summary>
        /// Writes to the file log
        /// </summary>
        /// <param name="LogText">Text to log</param>
        public static void WriteLog(string LogText)
        {
            try
            {
                string strFileName = System.AppDomain.CurrentDomain.FriendlyName + ".log.txt";
                using (StreamWriter w = File.AppendText(strFileName))
                {
                    w.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + LogText);
                }
            }
            catch
            {
                MessageBox.Show("Error! Unable to write to log file. This tool will now exit.");
                CommitSuicide();
            }
        }

        /// <summary>
        /// Violently kills the application.
        /// </summary>
        public static void CommitSuicide()
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }
    }

    public static class Extensions
    {

        /// <summary>
        /// Formats a DateTime object for SQL Insert (yyyy-MM-dd HH:mm:ss)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ToSQLDatetime(this DateTime dt)
        {
            return dt.ToString("yyyy-MM-dd HH:mm:ss");
        }
    }
}
