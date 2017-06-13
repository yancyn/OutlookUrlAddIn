﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookUrlAddIn
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Application application = this.Application;

            this.inspectors = this.Application.Inspectors;
            this.inspectors.NewInspector += inspectors_NewInspector;

            Outlook.Inspector activeInspector = application.ActiveInspector();
            if (activeInspector != null)
            {
                MessageBox.Show("Active inspector: " + activeInspector.Caption);
            }

            // Get the Explorer objects
            Outlook.Explorers explorers = application.Explorers;
            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                MessageBox.Show("Active explorer: " + activeExplorer.Caption);
            }
        }

        private void inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "Generated by Outlook Add-In";
                    mailItem.Body = "Hi!";
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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
