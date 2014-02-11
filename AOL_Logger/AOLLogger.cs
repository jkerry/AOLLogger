using log4net;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AOL_Logger
{
    public partial class AOLLogger
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
        private int ms_between_aggregations;
        private int HoursToKeep;
        private string folderPrefix;

        private static readonly ILog log = LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            getConfigValues();
            log4net.Config.XmlConfigurator.Configure();
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
            ms_between_aggregations = Properties.Settings.Default.ms_between_aggregations;
            folderPrefix = Properties.Settings.Default.folderPrefix;
            HoursToKeep = Properties.Settings.Default.hours_to_keep;
        }

        private void InstantiateSqlPreparedStatement()
        {
            preparedInsert = new SqlCommand("INSERT INTO AOLLogger.dbo.LoggedItems (Folder, Subject, Body, RecieptTime) VALUES (@Folder, @Subject, @Body, @RecieptTime)", myConnection);
            preparedDelete = new SqlCommand("DELETE FROM AOLLogger.dbo.LoggedItems WHERE RecieptTime < @DeleteTime", myConnection);

            FolderParam = new SqlParameter("@Folder", SqlDbType.NVarChar, 50);
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

        private void HandleFolder(SqlTransaction txn, Outlook.Folder topLevelFolder)
        {
            foreach (Outlook.Folder fol in topLevelFolder.Folders)
            {
                log.Info(string.Format("Processing {0}", fol.Name));
                //regex match
                if (Regex.IsMatch(fol.Name, String.Format(@"{0}.*", folderPrefix)))
                { //If this folder has an AOL_prefix then log the mail items contained and then delete them
                    DeleteAndLog(txn, fol);
                }
                HandleFolder(txn, fol);
            }
        }

        private void DeleteAndLog(SqlTransaction txn, Outlook.Folder folder)
        {
            int itemsLoggedInfolder = 0;
            string sourceFolder = folder.FolderPath;
            Stack<Outlook.MailItem> mailItems = new Stack<Outlook.MailItem>();
            foreach (Outlook.MailItem mail in folder.Items)
            {
                try
                {
                    string subject = mail.Subject;
                    string body = mail.Body;
                    DateTime timestamp = mail.ReceivedTime;
                    itemsLoggedInfolder += LogEmail(txn, sourceFolder, subject, body, timestamp);
                    mailItems.Push(mail);
                }
                catch (Exception e)
                {
                    LogException(e);
                }
            }
            //Delete in reverse order
            log.Info(string.Format("Deleting {0} items", mailItems.Count));
            foreach (Outlook.MailItem mail in mailItems)
            {
                mail.Delete();
            }
            log.Info(string.Format("Records inserted in {0}: {1}", folder.Name, itemsLoggedInfolder));
        }

        private void LogException(Exception e)
        {
            log.Error(string.Format("{0}\n{1}", e.Message, e.StackTrace));
        }

        private int LogEmail(SqlTransaction txn, string SourceFolder, string Subject, string Body, DateTime RecieptTime)
        {
            var logs = Regex.Split(Body, @"^\s*$[\r\n]*", RegexOptions.Multiline);
            int entriesAdded = 0;
            foreach (var log in logs)
            {
                if (!String.IsNullOrWhiteSpace(log))
                {
                    preparedInsert.Parameters[0].Value = SourceFolder;
                    preparedInsert.Parameters[1].Value = Subject;
                    preparedInsert.Parameters[2].Value = log;
                    preparedInsert.Parameters[3].Value = RecieptTime;
                    preparedInsert.Transaction = txn;
                    preparedInsert.Prepare();
                    entriesAdded += preparedInsert.ExecuteNonQuery();
                }
            }
            return entriesAdded;
        }

        private void DeleteOldLoggedEmails(SqlTransaction txn, DateTime DeletionDay)
        {
            preparedDelete.Parameters[0].Value = DeletionDay;
            preparedDelete.Transaction = txn;
            preparedDelete.Prepare();
            log.Info(string.Format("{0} Rows dropped from the database.", preparedDelete.ExecuteNonQuery()));
        }

        #endregion Private Helpers, Plugin logic

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

            myConnection = new SqlConnection(connectionString);

            while (true)
            {
                try
                {
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(ms_between_aggregations);

                        myConnection.Open();
                        var txn = myConnection.BeginTransaction();
                        InstantiateSqlPreparedStatement();

                        HandleFolder(txn, thisOutlookApplication.Session.DefaultStore.GetRootFolder() as Outlook.Folder);

                        if (HoursToKeep != 0) DeleteOldLoggedEmails(txn, DateTime.Now.AddHours(-1 * HoursToKeep));
                        txn.Commit();
                        myConnection.Close();
                    }
                }
                catch (Exception exception)
                {
                    log.Error("Worker thread level exception: " + exception.Message + "\n" + exception.StackTrace);
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

        #endregion Bankground Worker Job Logic

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

        #endregion VSTO generated code
    }
}