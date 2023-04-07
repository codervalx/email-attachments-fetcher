using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Timer = System.Timers.Timer;
using System.Timers;
using System.IO;
using System.Text.RegularExpressions;

namespace EmailAttachmentsFetcher
{
    public partial class Form1 : Form
    {
        private EmailFetcher emailFetcher;
        private bool started = false;

        public Form1()
        {
            InitializeComponent();
            tboxStatus.Text = "not running";
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (btnStart.Text == "START")
            {
                DialogResult result = folderBrowserDialog1.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string folderPath = folderBrowserDialog1.SelectedPath;
                    started = true;
                    emailFetcher = new EmailFetcher(folderPath);
                    Hide();
                    notifyIcon1.Visible = true;
                    emailFetcher.StartFetch();
                    btnStart.Text = "END";
                    tboxStatus.Text = "running";
                    //Console.WriteLine("Started...");
                }
            }
            else
            {
                started = false;
                emailFetcher.EndFetch();
                btnStart.Text = "START";
                tboxStatus.Text = "not running";
                //Console.WriteLine("Ended...");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (started)
            {
                emailFetcher.EndFetch();
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            Show();
            notifyIcon1.Visible = false;
            WindowState = FormWindowState.Normal;
        }
    }

    public class EmailFetcher
    {
        private Timer fetchTimer;
        private readonly string folderPath;
        private readonly string rootFolder = "Fetched Email Attachments";
        private Outlook.Application outlookApplication;
        private Outlook.NameSpace outlookNamespace;
        private Outlook.MAPIFolder inbox;
        private Outlook.Items items;
        private List<Outlook.MailItem> unreadMessagesWithAttachments;

        public EmailFetcher(string folderPath)
        {
            this.folderPath = folderPath;
            this.unreadMessagesWithAttachments = new List<Outlook.MailItem>();
            
            fetchTimer = new Timer();
            fetchTimer.Interval = 5000;
            fetchTimer.Elapsed += FetchTimer_Elapsed;
        }

        public void StartFetch()
        {
            fetchTimer.Start();
        }

        public void EndFetch()
        {
            fetchTimer.Stop();
            fetchTimer.Dispose();

            Marshal.ReleaseComObject(items);
            Marshal.ReleaseComObject(inbox);
            Marshal.ReleaseComObject(outlookNamespace);
            Marshal.ReleaseComObject(outlookApplication);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private async void FetchTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            await FetchEmailWithAttachmentsAsync(folderPath);
        }

        private async Task FetchEmailWithAttachmentsAsync(string folderPath)
        {
            try
            {
                string baseDir = Path.Combine(folderPath, rootFolder);
                if (!Directory.Exists(baseDir))
                {
                    Directory.CreateDirectory(baseDir);
                }

                outlookApplication = new Outlook.Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                items = inbox.Items.Restrict("[UnRead]=true");

                var tasks = new List<Task>();
                foreach (Outlook.MailItem item in items)
                {
                    tasks.Add(Task.Run(() =>
                    {
                        string itemSubjectFolder = "";
                        if (item.Attachments.Count > 0)
                        {
                            unreadMessagesWithAttachments.Add(item);

                            foreach (Outlook.Attachment attachment in item.Attachments)
                            {
                                string itemDateFolder = Path.Combine(baseDir, item.ReceivedTime.ToString("MM-dd-yyyy"));
                                itemSubjectFolder = Path.Combine(itemDateFolder, Regex.Replace(item.Subject, @"[^a-zA-Z0-9\s]+", ""));

                                if (!Directory.Exists(itemDateFolder))
                                {
                                    Directory.CreateDirectory(itemDateFolder);
                                }
                                if (!Directory.Exists(itemSubjectFolder))
                                {
                                    Directory.CreateDirectory(itemSubjectFolder);
                                }

                                attachment.SaveAsFile(Path.Combine(itemSubjectFolder, attachment.FileName));
                            }
                            item.SaveAs(Path.Combine(itemSubjectFolder, "mailmsg.msg"));
                        }
                    }));
                }
                await Task.WhenAll(tasks.ToArray());

                foreach (Outlook.MailItem item in unreadMessagesWithAttachments)
                {
                    item.UnRead = false;
                    item.Save();
                }

                unreadMessagesWithAttachments.Clear();

                await Task.Delay(1000);
                //Console.WriteLine("Done...");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occured. {ex}");
            }
        }
    }
}
