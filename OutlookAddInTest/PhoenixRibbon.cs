﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookAddInTest
{
	[ComVisible(true)]

	public class PhoenixRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public PhoenixRibbon()
		{
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("OutlookAddInTest.Ribbon1.xml");
		}

		#endregion

		#region Ribbon Callbacks
		//Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

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
				ArchiveForm form = new ArchiveForm(item);
				form.Show();
			}
		}

		private Outlook.MailItem getMailItem()
		{
			try
			{
				Object explorer = Globals.PhoenixPlugin.Application.ActiveWindow().Selection[1];
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
