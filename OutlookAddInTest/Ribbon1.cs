using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddInTest
{
	[ComVisible(true)]

	public class Ribbon1 : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;
		private Outlook.MailItem mailItem;

		public Ribbon1()
		{
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("OutlookAddInTest.Ribbon1.xml");
		}

		#endregion

		#region Ribbon Callbacks

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		#endregion

		public void OnTextButton(Office.IRibbonControl control)
		{
			Outlook.MailItem item = getMailItem();
			if (item != null)
			{
				FormTest form = new FormTest(item);
				form.Show();
			}
		}

		private Outlook.MailItem getMailItem()
		{
			try
			{
				Object explorer = Globals.ThisAddIn.Application.ActiveWindow().Selection[1];
				if (explorer is Outlook.MailItem)
				{
					return (Outlook.MailItem)explorer;
				}
				System.Diagnostics.Debug.WriteLine("The item you selected was not an email item.");
				return null;
			}
			catch (System.Runtime.InteropServices.COMException)
			{
				System.Windows.Forms.MessageBox.Show("You must select an Email first.");
				return null;
			}
		}

		public void OnRemoveButton(Office.IRibbonControl control)
		{
			Outlook.MailItem mailItem = getMailItem();
			if (mailItem != null)
			{
				if (mailItem.Categories != null)
				{
					if (mailItem.Categories.Contains("Phoenix archived"))
					{
						mailItem.Categories = mailItem.Categories.Replace("Phoenix archived", "");
						if (mailItem.Categories != null)
						{
							mailItem.Categories = mailItem.Categories.Replace(",,", ",");
						}
					}
				}
				mailItem.Save();
			}
		}

		#region Helpers

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
