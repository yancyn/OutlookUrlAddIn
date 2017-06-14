using System;
using System.Collections.Generic;
using System.IO;
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
                //MessageBox.Show("Active explorer: " + activeExplorer.Caption);

                string fileName = DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".txt";
                fileName = Path.GetTempPath() + fileName;

                StreamWriter writer = new StreamWriter(fileName);
                StringBuilder builder = new StringBuilder();

                foreach (Outlook.MailItem mail in activeExplorer.CurrentFolder.Items)
                {
                    // Extract url from email body
                    string[] urls = Utils.ExtractUrl(mail.Body);
                    if (urls.Length > 0)
                    {
                        foreach (string url in urls)
                            builder.AppendLine(url);
                    }
                    System.Diagnostics.Debug.WriteLine(mail.Subject);
                }

                // write result to text file
                writer.Write(builder.ToString());
                writer.Close();

                // open the text file
                System.Diagnostics.Process.Start(fileName);
            }
        }
    }
}