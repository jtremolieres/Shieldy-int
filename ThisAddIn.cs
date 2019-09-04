using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System.Threading;

namespace NotifySecurity
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

                CreateRibbonExtensibilityObject();
                #region declaration of the new event

                currentExplorer = this.Application.ActiveExplorer();
                if (currentExplorer == null) return;
                currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler
                    (CurrentExplorer_Event);
                #endregion

        }


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
          {
              return new Ribbon1();
          }




        #region action when a new mail is selected
        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;

            bool SelectedObjectIsMail = false;
            
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];


                    if (selObject is Outlook.MailItem)
                    {
                        SelectedObjectIsMail = true;
                       
                    }

                }
               

                ThisRibbonCollection ribbonCollection =Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
                if (SelectedObjectIsMail)
                {
                   


                }
            }
            catch (Exception )
            {
              
            }
          
        }
        #endregion

        #region action when a new mail is created
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }
        #endregion

        #region DO NOT TOUCH
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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

        #endregion

    }
}
