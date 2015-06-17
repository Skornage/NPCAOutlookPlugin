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
        DataTable dt = new DataTable();
        BindingSource bs = new BindingSource();
		public FormTest(Outlook.MailItem mailItem)
		{
			InitializeComponent();
            DisplayAccountInformation(Globals.ThisAddIn.Application);
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

		private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void Cancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Ok_Click(object sender, EventArgs e)
		{
            DataGridViewRow row = dataGridView1.CurrentRow;
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Directory.GetCurrentDirectory() + @"\Visual Studio 2013\Projects\NPCAOutlookPlugin\Output\Selections.txt", true))
            {
                  file.WriteLine(string.Format("id: {0}, type: {1}, name: {2}, email: {3}, info: {4}. Time: {5}",
                  row.Cells["id"].Value, row.Cells["type"].Value, row.Cells["name"].Value, row.Cells["email"].Value,
                  row.Cells["info"].Value, DateTime.Now));
            }
			this.Close();
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
            if (comboBox1.SelectedIndex == 1)
            {
                dt.DefaultView.RowFilter = "type = 'Account'";
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                dt.DefaultView.RowFilter = "type = 'Contact'";
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

            JObject dynJson = JsonTest.getJsonObjectTwo();
            foreach (KeyValuePair<String, JToken> item in dynJson)
            {
                String[] line = {(String) item.Value["id"], (String) item.Value["type"],
					(String) item.Value["name"], (String) item.Value["email"], (String) item.Value["info"]};
                dt.Rows.Add(line);
            }
            bs.DataSource = dt;
            dataGridView1.DataSource = bs;
        }

        private void FormTest_Load(object sender, EventArgs e)
        {

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
	


        public static void DisplayAccountInformation(Outlook.Application application)
        {

            // The Namespace Object (Session) has a collection of accounts.
            Outlook.Accounts accounts = application.Session.Accounts;

            // Concatenate a message with information about all accounts.
            StringBuilder builder = new StringBuilder();

            // Loop over all accounts and print detail account information.
            // All properties of the Account object are read-only.
            foreach (Outlook.Account account in accounts)
            {

                // The DisplayName property represents the friendly name of the account.
                builder.AppendFormat("DisplayName: {0}\n", account.DisplayName);

                // The UserName property provides an account-based context to determine identity.
                builder.AppendFormat("UserName: {0}\n", account.UserName);

                // The SmtpAddress property provides the SMTP address for the account.
                builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress);

                // The AccountType property indicates the type of the account.
                builder.Append("AccountType: ");
                switch (account.AccountType)
                {

                    case Outlook.OlAccountType.olExchange:
                        builder.AppendLine("Exchange");
                        break;

                    case Outlook.OlAccountType.olHttp:
                        builder.AppendLine("Http");
                        break;

                    case Outlook.OlAccountType.olImap:
                        builder.AppendLine("Imap");
                        break;

                    case Outlook.OlAccountType.olOtherAccount:
                        builder.AppendLine("Other");
                        break;

                    case Outlook.OlAccountType.olPop3:
                        builder.AppendLine("Pop3");
                        break;
                }

                builder.AppendLine();
            }

            // Display the account information.
            System.Windows.Forms.MessageBox.Show(builder.ToString());
        }
    }
}