using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInTest
{
	public partial class FormTest : Form
	{
		public FormTest(Outlook.MailItem mailItem)
		{
			InitializeComponent();
			populateListBox();
			comboBox1.SelectedIndex = 0;
		}

		private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void Cancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Ok_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void populateListBox()
		{
			dynamic dynJson = JsonTest.getJsonObject().DeserializeObject(json);
			foreach (var item in dynJson)
			{
				Console.WriteLine("{0} {1} {2} {3}\n", item.id, item.displayName, 
				item.slug, item.imageUrl);
			}
		}
	}
}
