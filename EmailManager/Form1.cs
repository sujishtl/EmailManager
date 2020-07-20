using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
//For Outlook 2010, you'll need to add reference to Microsoft.Office.Interop.Outlook version 14.0.0.0 else go with the latest
using System.IO;
namespace EmailManager
{
    public partial class Form1 : Form
    {
        string path = @"D:\log.txt";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private List<string> ReadMail()
        {
            List<string> mailList = new List<string>();
            try
            {
                var outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
                var outlookNameSpace = outlookApplication.GetNamespace("MAPI");
                //  var inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders["Optus"];
                var inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                Log("FolderName:" + inboxFolder.FullFolderPath.ToString());
                var mailItems = inboxFolder.Items;
                int count = mailItems.Count;//3122
                int index = 0;
                List<string> objectTypes = new List<string>();

                //foreach (Object o in mailItems)
                //{
                //    if (!objectTypes.Contains(o.GetType().ToString()))
                //    {
                //        objectTypes.Add(o.GetType().ToString());
                //        Log(o.GetType().ToString());
                //    }
                //}


                //foreach ( Object o in mailItems)
                //{
                //    //try
                //    //{
                //    //    var item = (MailItem)o;
                //    //}
                //    //catch(System.Exception ex)
                //    //{

                //    //}


                //  if(typeof(MailItem)==o.GetType())
                //    {

                //    }
                //  else
                //    {
                //        mailItems.Remove(index);
                //       Log("Removed Line:" + index);
                //    }
                //    index = index + 1;
                //}

                //  foreach (MailItem item in mailItems)
                for (int x = 1; x < count; x++)
                {

                    //if (mailItems[x].SentOn > DateTime.Now.AddMinutes(-9000))
                    //{

                        mailList.Add(mailItems[x].SenderEmailAddress + "," + mailItems[x].SentOn + "," + mailItems[x].Subject);

                    //}
                }


            }
            catch (System.Exception ex)
            {
                statusLabel.Text = "Exception Occured" + ex.Message;
                Log(ex.Message);
            }

            statusLabel.Text = "Getting items completed";
            return mailList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var csvPath = @"D:\csvMails"+DateTime.Now.ToString("ddMMyyHHmmss")+".csv";
            var itemsList = ReadMail();
            using (var file = File.CreateText(csvPath))
            {
                foreach (string item in itemsList)
                {
                    listBox1.Items.Add(item);
                    file.WriteLine(item);
                }
            }
        }

        public void Log(string content)
        {
            if (!File.Exists(path))
            {
                // Create a file to write to.

                File.WriteAllText(path, content + Environment.NewLine);
            }
            else
                File.AppendAllText(path, content + Environment.NewLine);
        }
    }
}
