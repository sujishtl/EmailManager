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
namespace EmailManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            var outlookNameSpace = outlookApplication.GetNamespace("MAPI");
            var inboxFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Folders["Optus"]; ;
            var mailItems = inboxFolder.Items;

            foreach( MailItem item in mailItems)
            {
                if(item.SentOn>DateTime.Now.AddMinutes(-1200))
                {
                    listBox1.Items.Add(item.Subject);
                    
                }
            }

        }
    }
}
