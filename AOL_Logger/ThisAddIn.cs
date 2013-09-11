using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.IO;

namespace AOL_Logger
{
    public partial class ThisAddIn
    {
        private BackgroundWorker bgw;
        private SqlConnection myConnection;
        private SqlCommand preparedInsert;
        private SqlCommand preparedDelete;
        private SqlParameter DeleteTimeParam;
        private SqlParameter FolderParam;
        private SqlParameter SubjectParam;
        private SqlParameter BodyParam;
        private SqlParameter RecieptTimeParam;

        private string connectionString;
        private int execution_hours;
        private int execution_minutes;
        private int DaysUntilAggregation;
        private int HoursToKeep;
        private string logFile;
        private string folderPrefix;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            getConfigValues();
            LaunchBackgroundLoggerService();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            KillBackgroundLoggerService();
        }

        #region Private Helpers, Plugin logic 

        private void getConfigValues()
        {
            connectionString = Properties.Settings.Default.ConnectionString;
            execution_hours = Properties.Settings.Default.execution_hours;
            execution_minutes = Properties.Settings.Default.execution_minutes;
            DaysUntilAggregation = Properties.Settings.Default.days_between_aggregations;
            HoursToKeep = Properties.Settings.Default.hours_to_keep;
            logFile = Properties.Settings.Default.logFile;
            folderPrefix = Properties.Settings.Default.folderPrefix;

        }

        private void InstantiateSqlPreparedStatement( )
        {
            preparedInsert = new SqlCommand("INSERT INTO AOLLogger.dbo.LoggedItems (Folder, Subject, Body, RecieptTime) VALUES (@Folder, @Subject, @Body, @RecieptTime)", myConnection);
            preparedDelete = new SqlCommand("DELETE FROM AOLLogger.dbo.LoggedItems WHERE RecieptTime < @DeleteTime", myConnection);

            FolderParam = new SqlParameter("@Folder",SqlDbType.NVarChar,50);
            SubjectParam = new SqlParameter("@Subject", SqlDbType.NVarChar, 256);
            BodyParam = new SqlParameter("@Body", SqlDbType.NText, 8000);
            RecieptTimeParam = new SqlParameter("@RecieptTime", SqlDbType.DateTime);
            DeleteTimeParam = new SqlParameter("@DeleteTime", SqlDbType.DateTime);

            preparedInsert.Parameters.Add(FolderParam);
            preparedInsert.Parameters.Add(SubjectParam);
            preparedInsert.Parameters.Add(BodyParam);
            preparedInsert.Parameters.Add(RecieptTimeParam);

            preparedDelete.Parameters.Add(DeleteTimeParam);
        }

        private void HandleFolder(Outlook.Folder topLevelFolder)
        {
            foreach (Outlook.Folder fol in topLevelFolder.Folders)
            {
                //regex match
                if (Regex.IsMatch(fol.Name, String.Format(@"{0}.*",folderPrefix)))
                { //If this folder has an AOL_prefix then log the mail items contained and then delete them
                    DeleteAndLog(fol);
                }
                HandleFolder(fol);
            }
        }


        private void DeleteAndLog(Outlook.Folder folder)
        {
            int itemsLoggedInfolder = 0;
            string sourceFolder = folder.FolderPath;
            Stack<Outlook.MailItem> mailItems = new Stack<Outlook.MailItem>();
            foreach( Outlook.MailItem mail in folder.Items ){
                try
                {
                    string subject = mail.Subject;
                    string body = mail.Body;
                    DateTime timestamp = mail.ReceivedTime;
                    itemsLoggedInfolder += LogEmail(sourceFolder, subject, body, timestamp);
                    mailItems.Push(mail);
                }
                catch (Exception e)
                {
                    LogException(e);
                }
            }
            //Delete in reverse order
            foreach(Outlook.MailItem mail in mailItems){
                mail.Delete();
            }
            System.Diagnostics.Debug.WriteLine("Records inserted in " + folder.Name + ": " + itemsLoggedInfolder);
        }

        private void LogException(Exception e)
        {


            try
            {
                StreamWriter sw = new StreamWriter(logFile);
                sw.WriteLine(DateTime.Now + "\n" + e.Message + "\n" + e.StackTrace);
                sw.Close();
            }
            catch (Exception ioException)
            {

            }

        }

        private int LogEmail(string SourceFolder, string Subject, string Body, DateTime RecieptTime)
        {

            preparedInsert.Parameters[0].Value = SourceFolder;
            preparedInsert.Parameters[1].Value = Subject;
            preparedInsert.Parameters[2].Value = Body;
            preparedInsert.Parameters[3].Value = RecieptTime;

            preparedInsert.Prepare();
            return preparedInsert.ExecuteNonQuery();
        }

        private void DeleteOldLoggedEmails(DateTime DeletionDay)
        {
            preparedDelete.Parameters[0].Value = DeletionDay;
            preparedDelete.Prepare();
            System.Diagnostics.Debug.WriteLine(preparedDelete.ExecuteNonQuery() + " Rows dropped from the database.");
        }

        #endregion


        #region Bankground Worker Job Logic

        private void LaunchBackgroundLoggerService()
        {
            bgw = new BackgroundWorker();
            bgw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            if (!bgw.IsBusy)
            {
                bgw.RunWorkerAsync();
            }

        }

        private void KillBackgroundLoggerService()
        {
            if (bgw != null) bgw.CancelAsync();
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Outlook.Application thisOutlookApplication = new Outlook.Application();

            DateTime executionTime = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, execution_hours, execution_minutes, 0);

            myConnection = new SqlConnection(connectionString);

            while (true)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    int milisecondsToWait = (int)((TimeSpan)(executionTime - DateTime.Now)).TotalMilliseconds;
                    while (milisecondsToWait <= 0)
                    {
                        executionTime = executionTime.AddDays(1);
                        milisecondsToWait = (int)((TimeSpan)(executionTime - DateTime.Now)).TotalMilliseconds;
                    }
                    System.Threading.Thread.Sleep(milisecondsToWait);
                    
                    myConnection.Open();

                    InstantiateSqlPreparedStatement();

                    HandleFolder(thisOutlookApplication.Session.DefaultStore.GetRootFolder() as Outlook.Folder);

                    DeleteOldLoggedEmails( executionTime.AddHours(-1*HoursToKeep) );

                    myConnection.Close();

                    
                }
            }
        }

        


        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Worker is continuing its work");
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((e.Cancelled == true))
            {
                System.Diagnostics.Debug.WriteLine("Canceled!");
            }

            else if (!(e.Error == null))
            {
                System.Diagnostics.Debug.WriteLine("Error: " + e.Error.Message);
            }

            else
            {
                System.Diagnostics.Debug.WriteLine("Done!");
            }
        }

        #endregion


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
