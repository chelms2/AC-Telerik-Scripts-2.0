using Telerik.TestStudio.Translators.Common;
using Telerik.TestingFramework.Controls.TelerikUI.Blazor;
using Telerik.TestingFramework.Controls.KendoUI.Angular;
using Telerik.TestingFramework.Controls.KendoUI;
using Telerik.WebAii.Controls.Html;
using Telerik.WebAii.Controls.Xaml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.Design;
using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;

// the below namespaces are needed for messagebox, csv handling and emailing results
// using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Mail;
using ArtOfTest.WebAii.Win32.Dialogs;

namespace ReportDistribution
{

    public class Report_Distribution_Script : BaseWebAiiTest
    {
        #region [ Dynamic Pages Reference ]

        private Pages _pages;

        /// <summary>
        /// Gets the Pages object that has references
        /// to all the elements, frames or regions
        /// in this project.
        /// </summary>
        public Pages Pages
        {
            get
            {
                if (_pages == null)
                {
                    _pages = new Pages(Manager.Current);
                }
                return _pages;
            }
        }

        #endregion
        
        // Add your test methods here...
        
        // Save location where the csv files are downloaded
        // Use below to run in server.
        // string saveLocation = @"\\utshare.local\cifs\Developers\za\SharedServices\Team\AccessControl\Telerik\";
        // use C:\Temp for local test
        string saveLocation = @"C:\Users\chelms\OneDrive - UT System\Desktop\Temp\";
    
        [CodedStep(@"New Coded Step")]
        public void Report_Distribution_Script_CodedStep()
        {
            // Click 'URL1Link' - download CSV Report link
            // Save the CSV file to saveLocation.
            // Keeping it in try catch bypass if any errors occur during the saveDialog is handled
            // That way the script will not crash.            
            try
            {                
                // Click 'URL1Link'
                //ActiveBrowser.Window.SetFocus();
                //Pages.ProcessMonitor4.FramePtModFrame2.URL1Link.ScrollToVisible(ArtOfTest.WebAii.Core.ScrollToVisibleType.ElementCenterAtWindowCenter);

                //Trigger the dialog.
                //string reportName = Pages.ProcessMonitor4.FramePtModFrame2.URL1Link.TextContent; 
                //Log.WriteLine(saveLocation + reportName);
                
                //DownloadDialogsHandler dialog = new DownloadDialogsHandler(Manager.ActiveBrowser, DialogButton.SAVE, saveLocation + reportName, Manager.Desktop); 
                //Manager.DialogMonitor.Start();
                //Pages.ProcessMonitor4.FramePtModFrame2.URL1Link.Click(false);
                //dialog.WaitUntilHandled(30000);
                           
            }
            catch (Exception e)
            {
                Log.WriteLine("The process at downloading: " + e.ToString());
                
            }
            finally {}
            
        }
        
        [CodedStep(@"New Coded Step")]
        public void Report_Distribution_Script_CodedStep1()
        {
            // declare messagebox title and exception message variables
            // If you want to debug the code, you may uncomment messageTitle and messagebox lines below.
            // Also uncomment the using System.Windows.Forms line at the top.
            
            //string messageTitle = "";
            string excMessages = "";
            string reportEnvironment = "";
            //MessageBoxButtons buttons = MessageBoxButtons.OK;
            
            // define the file lookup extention, this case .csv
            string pattern = "*.csv";
            var dirInfo = new DirectoryInfo(saveLocation);
            
                        
            
            // following line finds the latest csv file under dirInfo directory
            var LatestFile = (from f in dirInfo.GetFiles(pattern) orderby f.LastWriteTime descending select f).First();            
            //NativeWindow a = new NativeWindow();            
            //a.AssignHandle(ActiveBrowser.Window.Handle);
            
            // show the most recent file on a messagebox and keep it in log file
            //MessageBox.Show(a, LatestFile.ToString(), "Latest File", buttons, MessageBoxIcon.Information);
            Log.WriteLine("Latest file - last step : " + LatestFile.ToString());
            
            // Gmail Smtp settings. uncomment this if you need to test it.
            /*
            string gUser = "telerikauditresults@gmail.com";
            string gPass = "123Telerik";
            // email handling
            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(gUser, gPass),
                EnableSsl = true,
                
            };      
            */
            
            // SMTP configuration without credentials       
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient();
            SmtpServer.Host = "mail.utshare.utsystem.edu";
            SmtpServer.Port = 25;
            SmtpServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            
            //configure from-to e-mail addresses.
            mail.To.Add("chelms@utsystem.edu");
            mail.From = new MailAddress("fromTelerik@utsystem.edu");
            
            // If you want to debug the code, you may uncomment messageTitle and messagebox lines below.
            // Also uncomment the using System.Windows.Forms line at the top. 
            // delegate (kind of function under function in C#) to print exception messages - 
            var printMessages = new Action<string, string>((exc, dbName) => {
                // if there is no exceptions in the csv file, prompt a message and keep it in the log file
                if(String.IsNullOrEmpty(exc))
                {
                    //messageTitle = "Clean";                    
                    //MessageBox.Show(a, "No SEC- exceptions found!", messageTitle, buttons, MessageBoxIcon.Information); 
                    Log.WriteLine(exc);    
                    
                    // Below two lines are for gmail smtp. uncomment them if you need to use.
                    // smtpClient.Send(gUser, "ekopru@utsystem.edu", "SYSAudit Report for: " + LatestFile.ToString() + " -- " + dbName , "\nThe audit is clean. No exceptions found!" + excMessages);
                    // Log.WriteLine("Email sent.");
                    
                    //configure mail components.
                    mail.Subject = LatestFile.ToString() + " -- " + dbName ;
                    mail.IsBodyHtml = false;
                    mail.Body = "\nThe audit is clean. No exceptions found!";

                    // Send email, catch if there is any error.
                    try {
                        SmtpServer.Send(mail);
                        Log.WriteLine("Email sent without credentials");
                    }
                    catch (Exception ex) {
                        Log.WriteLine("Exception Message: " + ex.Message);
                        if (ex.InnerException != null)
                            Log.WriteLine("Exception Inner:   " + ex.InnerException);
                    }                    
                }
                
                // if there is/are exception(s) in the csv file, prompt a message, email the exceptions
                // keep them in the log file, finally throw an exception to fail the test if you want to. (optional)
                // At the moment it wont fail the test.
                else
                {
                    //messageTitle = "SEC- Errors found!";
                    //MessageBox.Show(a, exc, messageTitle, buttons, MessageBoxIcon.Warning); 
                    Log.WriteLine(exc);                                      
                    // Below two lines are for gmail smtp. uncomment them if you need to use.
                    // smtpClient.Send(gUser, "ekopru@utsystem.edu", "SYSAudit Report for: " + " -- " + dbName, "(SEC- Exceptions ---\n" + excMessages);
                    // Log.WriteLine("Email sent.");
                    
                    //configure mail components.
                    mail.Subject = LatestFile.ToString() + " -- " + dbName ;
                    mail.IsBodyHtml = false;
                    mail.Body = "Security Exceptions --- \n" + excMessages;
                    
                    // Send the email, catch if there is any error.
                    try {
                        SmtpServer.Send(mail);
                        Log.WriteLine("Email sent without credentials");
                    }
                    catch (Exception ex) {
                        Log.WriteLine("Exception Message: " + ex.Message);
                        if (ex.InnerException != null)
                            Log.WriteLine("Exception Inner:   " + ex.InnerException);
                    }  
                    // fail test
                    // throw new System.IO.FileNotFoundException("SEC- Errors found! ", exc.ToString());
                }               
            });  
            
            // CSV file handling happens below. Loops thru the CSV file and collects every cell that has a value starting with (SEC-
            // Also captures DB name, Run data and Run time.
            using (Microsoft.VisualBasic.FileIO.TextFieldParser parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(saveLocation + LatestFile.ToString()))
            {                
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = false;

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                                        
                    for(int i = 0; i < fields.Length; i++)
                    
                    {
                        if(fields[i].StartsWith("(SEC-"))
                        {
                            excMessages = excMessages + fields[i] + "\n\n";                            
                        }
                                                
                        else if(fields[i].StartsWith("Database Name"))
                        {
                            reportEnvironment += fields[i] + " " + fields[i+1] + "  --  " + fields[i+2] + "  " + fields[i+3];                            
                        }
                        else if(fields[i].StartsWith("Run Time"))
                        {
                            reportEnvironment += "  --  " + fields[i] + "  " + fields[i+1];                            
                        }
                    }                               
                }
                printMessages(excMessages, reportEnvironment); 
            }
        }
    }
}
                
        
    

