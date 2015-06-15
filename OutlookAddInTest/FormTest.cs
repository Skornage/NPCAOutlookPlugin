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
            string test = JsonTest.getJsonObject();
            
            Newtonsoft.Json.Linq.JToken json = Newtonsoft.Json.Linq.JToken.Parse(test);

			foreach (var item in json.First.First)
			{
                String[] line = {(String)item["id"], (String)item["type"],
                (String)item["name"], (String)item["email"], (String)item["info"]};
                //String line = String.Format("{0}, {1}, {2}, {3}, {4}", item["id"], item["type"],
                //item["name"], item["email"], item["info"]);
                listView1.Items.Add(new ListViewItem(line));
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
