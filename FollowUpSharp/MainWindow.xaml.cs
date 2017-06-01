using System;
using System.Collections.Generic;
using System.Windows;
using System.ComponentModel;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.Threading;

namespace FollowUpSharp
{
    public partial class MainWindow : Window
    {
        private List<string> attachedFiles = new List<string>(); // Holds user-specified filepaths
        private BackgroundWorker worker = new BackgroundWorker(); // New instance for UI updating
        
        public MainWindow()
        {
            InitializeComponent();

            // Set up the BackgroundWorker
            worker.WorkerReportsProgress = true;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private void SendEmails_Click(object sender, RoutedEventArgs e)
        {
            SendEmails.IsEnabled = false;
            object[] parameters = { Queries.Text, attachedFiles };
            worker.RunWorkerAsync(parameters);
            // .NET 4.0
            // new Tuple<string, List<string>>(Queries.Text, attachedFiles)
        }
        /*
         * Will remove later
         * 
         * Jimmy explained how and why I was stuck with the BackgroundWorker:
         * It wasn't a matter of not understanding how threads work or race conditions,
         * but how the type system worked. I was running into the issue of having things
         * I wanted to pass to a method in Worker_DoWork, but BackgroundWorker's
         * parameter is one object. This means it can be literally anything (since
         * everything derives from System.Object) which is why there are solutions
         * to this online that simply declare an object array, shove any required
         * paramters in said array, and pass it through via BackgroundWorker's parameter.
         * We could then unpack this array by declaring variables in the Worker_DoWork
         * method to match our parameters and assign the respective contents of the array
         * to those parameters.
         * 
         * What Jimmy suggested was in essence the same idea but pack everything in a tuple
         * and do everything I would for an object array but with a tuple since I only had
         * a max of two variables to deal with.
         * 
         * And, it worked! This is why having someone to actually talk to who's better than you
         * is so useful, someone can sanity check you and tell you what you're doing wrong.
         * 
         * Downgrading the project to .NET Framework 3.5 to support Windows 7 without hassle
         * required switching from a tuple to an object array anyway.
         */

        #region BackgroundWorker for Progress Bar
        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // .NET 4.0
            // var parameters = (Tuple<string, List<string>>)e.Argument;
            // Unpack the contents of the object array into two variables and
            // then use them as necessary
            var parameters = (object[])e.Argument;
            string queryText = (string)parameters[0];
            List<string> files = (List<string>)parameters[1];
            
            if (attachedFiles.Count == 0)
            {
                FollowUp.SendFollowUps(queryText, worker.ReportProgress);
            }
            else
            {
                FollowUp.SendFollowUps(queryText, worker.ReportProgress, files);
                // Clear the GUI and then clear the list
                FileList.Items.Clear();
                attachedFiles = null;
            }

        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progress.Value = e.ProgressPercentage;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SendEmails.IsEnabled = true;
            Thread.Sleep(1000);
            Progress.Value = 0;
            MessageBox.Show("Follow ups sent!");
        }
        #endregion

        #region GUI Functionality
        /// <summary>
        /// Clicking "Access Excel Records" in the File menu will open Explorer to the archive for
        /// all of the Quote Follow Ups sent via the program
        /// </summary>
        private void Excel_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(
                "explorer.exe",
                $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\FollowUpSharp\Quote Follow Ups Archive\"
                );
        }

        /// <summary>
        /// You know what it is
        /// </summary>
        private void QuitProgram_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// View error logs caught by program
        /// </summary>
        private void ErrorLogs_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\ErrorLog.txt"))
            {
                Process.Start("notepad.exe", $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\FollowUpSharp\ErrorLog.txt");
            }
            else
            {
                MessageBox.Show("No application-halting errors have been recorded!");
            }
            
        }

        /// <summary>
        /// Opens the help file
        /// </summary>
        private void Help_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("iexplore.exe", $@"C:\Program Files (x86)\WKFC Auto Follow Up\help\helphome.html");
        }

        /// <summary>
        /// The functionality of the "browse" button for attaching files to the email.
        /// </summary>
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            ofd.Multiselect = true;
            ofd.ShowDialog();
            string[] SelectedFiles = ofd.FileNames;
            foreach (string file in SelectedFiles)
            {
                attachedFiles.Add(file);
                int slash = file.LastIndexOf("\\");
                string filename = file.Substring(slash + 1);
                FileList.Items.Add(filename.Trim());
            }
        }

        /// <summary>
        /// Just in case the user wants to get rid of the file(s) they were going to attach,
        /// allow the users to clear the attachment list
        /// </summary>
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            FileList.Items.Clear();
            attachedFiles.Clear();
        }

        #endregion


    }
}
