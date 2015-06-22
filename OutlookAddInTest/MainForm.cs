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
	public partial class MainForm : Form
	{
		private Outlook.MailItem mailItem;
        DataTable dt = new DataTable();
        BindingSource bs = new BindingSource();
		public MainForm(Outlook.MailItem mailItem)
		{
			this.mailItem = mailItem;
			InitializeComponent();
            populateDataGrid(0);
			comboBox1.SelectedIndex = 0;
            textBox1.TextChanged += new EventHandler(searchTextChanged);
		}

        private void searchTextChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != 0)
            {
                bs.Filter = string.Format("name LIKE '{0}%' AND type='{1}'", textBox1.Text, comboBox1.SelectedItem);
            }
            else
            {
                bs.Filter = string.Format("name LIKE '{0}%'", textBox1.Text);
            }         
        }

		private void Cancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Ok_Click(object sender, EventArgs e)
		{
			//DataGridViewRow row = dataGridView1.CurrentRow;
			//using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Directory.GetCurrentDirectory() + @"\Visual Studio 2013\Projects\NPCAOutlookPlugin\Output\Selections.txt", true))
			//{
			//	file.WriteLine(string.Format("id: {0}, type: {1}, name: {2}, email: {3}, info: {4}. Time: {5}",
			//	row.Cells["id"].Value, row.Cells["type"].Value, row.Cells["name"].Value, row.Cells["email"].Value,
			//	row.Cells["info"].Value, DateTime.Now));
			//}

            //API.Archive();
            mailItem.MessageClass = "IPM.Note.Phoenix";
            mailItem.UserProperties.Remove(1);
			mailItem.Save();
			this.Close();
		}

		private bool CategoryExists(string categoryName)
		{
			try
			{
				Outlook.Category category = mailItem.Application.Session.Categories[categoryName];
				return (category != null);
			}
			catch { return false; }
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
            if (comboBox1.SelectedIndex != 0)
            {
                dt.DefaultView.RowFilter = string.Format("type = '{0}' AND name LIKE '%{1}%'", comboBox1.SelectedItem, textBox1.Text);
            }
            else
            {
                dt.DefaultView.RowFilter = "";
            }
		}

        private void populateDataGrid(int index)
        {
            dt.Columns.Add("id", typeof(String));
            dt.Columns.Add("type", typeof(String));
            dt.Columns.Add("name", typeof(String));
            dt.Columns.Add("email", typeof(String));
            dt.Columns.Add("info", typeof(String));

            JObject dynJson = JsonGetter.getJsonObject();
            foreach (KeyValuePair<String, JToken> item in dynJson)
            {
                String[] line = {(String) item.Value["id"], (String) item.Value["type"],
					(String) item.Value["name"], (String) item.Value["email"], (String) item.Value["info"]};
                dt.Rows.Add(line);
            }
            bs.DataSource = dt;
            dataGridView1.DataSource = bs;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

		private void button3_Click(object sender, EventArgs e)
		{
            SearchForm searchForm = new SearchForm(this);
            searchForm.Show();
		}

        public void advancedSearch(string id, string type, string name, string email, string info)
        {
            comboBox1.SelectedIndex = 0;
            bs.Filter = string.Format("id LIKE '%{0}%' AND type LIKE '%{1}%' AND name LIKE '%{2}%' AND email LIKE '%{3}%' AND info LIKE '%{4}%'",
                                       id, type, name, email, info);
        }

		private void FormTest_Load(object sender, EventArgs e)
		{

		}

		private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

        //public static void DisplayAccountInformation(Outlook.Application application)
        //{
        //    Outlook.Accounts accounts = application.Session.Accounts;
        //    StringBuilder builder = new StringBuilder();

        //    foreach (Outlook.Account account in accounts)
        //    {
        //        builder.AppendFormat("DisplayName: {0}\n", account.DisplayName);
        //        builder.AppendFormat("UserName: {0}\n", account.UserName);
        //        builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress);
        //        builder.Append("AccountType: ");
        //        builder.AppendLine();
        //    }

        //    // Display the account information.
        //    System.Windows.Forms.MessageBox.Show(builder.ToString());
        //}
    }
}