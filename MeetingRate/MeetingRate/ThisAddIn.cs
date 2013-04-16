using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace MeetingRate
{
    public partial class ThisAddIn
    {

        Outlook.Inspectors inspectors;
        Outlook.AppointmentItem appointment;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            ((Outlook.ApplicationEvents_Event)this.Application).ItemSend += ThisAddIn_ItemSend;
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            if (Inspector.CurrentItem is Outlook.AppointmentItem)
                appointment = Inspector.CurrentItem as Outlook.AppointmentItem;
            else
                appointment = null;
        }

        void ThisAddIn_ItemSend(object Item, ref bool Cancel)
        {
            if (appointment != null)
                    {
                        if (appointment.EntryID == null)
                        {
                            if (appointment.Recipients != null)
                            {
                                string importantMessage = "Minimum cost of this meeting is " + appointment.Recipients.Count * 30 + " EUR. Let's be efficient!" + Environment.NewLine + Environment.NewLine;

                                string messageBoxText = "Meeting you are going to organize will cost at least: " + appointment.Recipients.Count * 30 + " EUR" + Environment.NewLine
                                    + Environment.NewLine
                                    + "Add this message to body?";

                                string caption = "Waste stopper!";
                                
                                MessageBoxButtons button = MessageBoxButtons.YesNoCancel;

                                DialogResult result = MessageBox.Show(messageBoxText, caption, button);

                                if (result == DialogResult.Yes)
                                {
                                    if (appointment.Body != null)
                                        appointment.Body = appointment.Body.Insert(0, importantMessage);
                                    else
                                        appointment.Body = importantMessage;
                                }
                                if (result == DialogResult.Cancel)
                                {
                                    Cancel = true;
                                }
                            }
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
