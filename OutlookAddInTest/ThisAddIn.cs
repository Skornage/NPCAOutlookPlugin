using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddInTest
{
    public partial class ThisAddIn
    {
		private Outlook.NameSpace outlookNameSpace;
		private Outlook.MAPIFolder inbox;
		private Outlook.Items items;
		private Outlook.MailItem selectedItem;
		private Outlook.MailItem responseItem;
		private Outlook.Explorer currentExplorer = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			this.currentExplorer = this.Application.ActiveExplorer();
            this.currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Custom_CurrentExplorer_Event);

			outlookNameSpace = this.Application.GetNamespace("MAPI");
			inbox = outlookNameSpace.GetDefaultFolder(
					Microsoft.Office.Interop.Outlook.
					OlDefaultFolders.olFolderInbox);

			items = inbox.Items;
			items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

		private void Custom_CurrentExplorer_Event()
		{
			try
			{
				if (this.Application.ActiveExplorer().Selection.Count > 0)
				{
                    Object selObject = this.currentExplorer.Selection[1]; // this.Application.ActiveExplorer().Selection[1];
					if (selObject is Outlook.MailItem)
					{
						this.selectedItem = (selObject as Outlook.MailItem);                    
                        if (this.selectedItem.UserProperties.Find("hasBeenSelected") == null)
                        {
                            ((Outlook.ItemEvents_10_Event)this.selectedItem).Reply += OnReply;
                        }
                        this.selectedItem.UserProperties.Add("hasBeenSelected", Outlook.OlUserPropertyType.olInteger);
                        this.selectedItem.Save();
					}
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void OnReply(object Response, ref bool Cancel)
		{
			this.responseItem = (Outlook.MailItem)Response;
            ((Outlook.ItemEvents_10_Event)this.responseItem).Send += new Outlook.ItemEvents_10_SendEventHandler(OnSend);
		}

		public void OnSend(ref bool Cancel)
		{
			if (this.selectedItem.Categories != null) 
			{
				if (this.selectedItem.Categories.Contains("Phoenix archived")) 
				{
                    this.responseItem.Categories += "Phoenix archived";
                    this.responseItem.Save();
				}
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("still in the event at least");
            }
		}

		void items_ItemAdd(object Item)
		{
			// if (APICall.isArchived(Item) 
			// {
				// APICAll.archive(Item);
				((Outlook.MailItem) Item).Categories += "Phoenix archived";
				((Outlook.MailItem)Item).Save();
			// }
		}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return new Ribbon1();
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
