using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookUrlAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application application = Globals.ThisAddIn.Application;

            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                MessageBox.Show("Active explorer: " + activeExplorer.Caption);
                foreach (Outlook.MailItem mail in activeExplorer.CurrentFolder.Items)
                {
                    System.Diagnostics.Debug.WriteLine(mail.Subject);
                }
            }
        }
    }
}
