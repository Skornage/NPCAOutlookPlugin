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
using Newtonsoft.Json.Linq;

namespace OutlookAddInTest
{
	public partial class FormTest : Form
	{
		public FormTest(Outlook.MailItem mailItem)
		{
			InitializeComponent();
			populateListBox(0);
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
			populateListBox(comboBox1.SelectedIndex);
		}

		private void populateListBox(int index)
		{
			JObject dynJson = JsonTest.getJsonObjectTwo();
			if (listView1.Items.Count > 0)
			{
				listView1.Items.Clear();
			}
			//listView1.Clear();
			if (index == 0 && listView1.Items.Count == 0)
			{
				foreach (KeyValuePair<String, JToken> item in dynJson)
				{
					String[] line = {(String) item.Value["id"], (String) item.Value["type"],
					(String) item.Value["name"], (String) item.Value["email"], (String) item.Value["info"]};
					listView1.Items.Add(new ListViewItem(line));
				}
			} else if (index == 1) {
				foreach (KeyValuePair<String, JToken> item in dynJson)
				{
					if (((String) item.Value["type"]).Equals("Account"))
					{
						String[] line = {(String) item.Value["id"], (String) item.Value["type"],
						(String) item.Value["name"], (String) item.Value["email"], (String) item.Value["info"]};
						listView1.Items.Add(new ListViewItem(line));
					}
				}
			} else {
				foreach (KeyValuePair<String, JToken> item in dynJson)
				{
					if (((String) item.Value["type"]).Equals("Contact"))
					{
						String[] line = {(String) item.Value["id"], (String) item.Value["type"],
						(String) item.Value["name"], (String) item.Value["email"], (String) item.Value["info"]};
						listView1.Items.Add(new ListViewItem(line));
					}
				}
			}
			listView1.Columns[0].Width = -2;
			listView1.Columns[1].Width = -2;
			listView1.Columns[2].Width = -2;
			listView1.Columns[3].Width = -2;
			listView1.Columns[4].Width = -2;
		}

        private void FormTest_Load(object sender, EventArgs e)
        {

        }
	}
}
