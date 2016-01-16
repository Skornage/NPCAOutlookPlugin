﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System.IO;
using System.Net;
using System.Net.Http;

namespace OutlookAddInTest
{
	public partial class MainForm : Form
	{
		private Outlook.MailItem mailItem;
        DataTable dt = new DataTable();
        BindingSource bs = new BindingSource();
		List<Result> results;
		private Outlook.ExchangeUser currentUser;

		public MainForm(Outlook.MailItem mailItem, PhoenixPlugin app)
		{
			this.currentUser = app.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
			this.mailItem = mailItem;
			InitializeComponent();
			initializeDataGrid();
			//comboBox1.SelectedIndex = 0;
			textBox1.setDelayedTextChangedTimerTickHandler(new EventHandler(HandleDelayedTextChangedTimerTick));
			searchLabel.Hide();
			//+= new EventHandler(searchTextChanged);
		}

		private void HandleDelayedTextChangedTimerTick(object sender, EventArgs e)
		{
			Timer timer = sender as Timer;
			timer.Stop();
			searchLabel.Show();
			results = textBox1.OnDelayedTextChanged(EventArgs.Empty);
			populateDataGrid(results);
			searchLabel.Hide();
		}

		//private void searchTextChanged(object sender, EventArgs e)
		//{
		//	if (comboBox1.SelectedIndex != 0)
		//	{
		//		bs.Filter = string.Format("name LIKE '{0}%' AND type='{1}'", textBox1.Text, comboBox1.SelectedItem);
		//	}
		//	else
		//	{
		//		bs.Filter = string.Format("name LIKE '{0}%'", textBox1.Text);
		//	}
		//}

		private void Cancel_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void Ok_Click(object sender, EventArgs e)
		{
			try
			{
				string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
				string PR_ATTACH_MIME_TAG = @"http://schemas.microsoft.com/mapi/proptag/0x370E001F";
				const string PR_ATTACH_DATA_BIN =
					"http://schemas.microsoft.com/mapi/proptag/0x37010102";

				var entityIdProperty = mailItem.UserProperties.Add("entityId",
					Outlook.OlUserPropertyType.olText, false, 1);

				//DataGridViewRow row = dataGridView1.CurrentRow;
				foreach (DataGridViewRow row in dataGridView1.SelectedRows)
				{
					String username = this.currentUser.Name;
					DateTime whenReceivedUtc = mailItem.ReceivedTime.ToUniversalTime();
					String fromDisplayName = mailItem.SenderName;
					var mailSender = mailItem.Sender;
					String fromEmailAddress = "";
					try
					{
						fromEmailAddress = mailSender.PropertyAccessor.GetProperty(
								PR_SMTP_ADDRESS) as string;
					}
					catch (Exception exc)
					{
						fromEmailAddress = "";
					}
					String subject = mailItem.Subject;
					String body = mailItem.HTMLBody;
					bool isBodyHtml = true;

					ArchiveEmailItem toArchive = new ArchiveEmailItem(username, whenReceivedUtc, fromDisplayName, fromEmailAddress,
						subject, body, isBodyHtml);

					//Get Attachments
					ArchiveEmailAttachment archAttachment;
					foreach (Outlook.Attachment attachment in mailItem.Attachments)
					{
						String fileName = attachment.FileName;
						//String mediaTypeName = attachment.PropertyAccessor.GetProperty(PR_ATTACH_MIME_TAG);
						byte[] content = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);
						archAttachment = new ArchiveEmailAttachment(fileName, content);
						toArchive.addAttachment(archAttachment);
					}

					archiveEmail(toArchive, row, entityIdProperty);
					mailItem.Save();
				}
				this.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
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

		//private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		//{
		//	if (comboBox1.SelectedIndex != 0)
		//	{
		//		dt.DefaultView.RowFilter = string.Format("type = '{0}' AND name LIKE '%{1}%'", comboBox1.SelectedItem, textBox1.Text);
		//	}
		//	else
		//	{
		//		dt.DefaultView.RowFilter = "";
		//	}
		//}

        private void initializeDataGrid()
        {
			dt = new DataTable();
			dt.Columns.Add("id", typeof(String));
			dt.Columns.Add("type", typeof(Bitmap));
            dt.Columns.Add("name", typeof(String));
			dt.Columns.Add("email", typeof(String));
			dt.Columns.Add("city", typeof(String));
			dt.Columns.Add("state", typeof(String));
			bs.DataSource = dt;
			dataGridView1.DataSource = bs;
        }

		private void populateDataGrid(List<Result> results) 
		{
			initializeDataGrid();
			foreach (Result item in results)
			{
				Bitmap type;
				if (item.isContact)
				{
					type = new Bitmap(AppDomain.CurrentDomain.BaseDirectory + @"..\..\res\fa-user.bmp");
				}
				else
				{
					type = new Bitmap(AppDomain.CurrentDomain.BaseDirectory + @"..\..\res\fa-industry.bmp");
				}

				Object[] line = {(String) item.idNumber.ToString(), (Bitmap) type,
					(String) item.name, (String) item.emailAddress,
								(String) item.city, (String) item.stateProvince};

				dt.Rows.Add(line);
			}
		}

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

		private void button3_Click(object sender, EventArgs e)
		{
			//SearchForm searchForm = new SearchForm(this);
			//searchForm.Show();
			results = textBox1.OnDelayedTextChanged(EventArgs.Empty);
			populateDataGrid(results);
		}

        public void advancedSearch(string id, string type, string name, string email)
        {
			//comboBox1.SelectedIndex = 0;
            bs.Filter = string.Format("id LIKE '%{0}%' AND type LIKE '%{1}%' AND name LIKE '%{2}%' AND email LIKE '%{3}%'",
                                       id, type, name, email);
        }

		private void FormTest_Load(object sender, EventArgs e)
		{

		}

		private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
		{

		}

		private void archiveEmail(ArchiveEmailItem email, DataGridViewRow row, Outlook.UserProperty entityIdProperty)
		{
			int id = int.Parse(row.Cells["id"].Value.ToString());
			String entityId = "";
			foreach (Result item in results)
			{
				if (item.idNumber == id)
				{
					entityId = item.id;
					break;
				}
			}
			String entryId = mailItem.EntryID;
			var jsonSerializerSettings = new JsonSerializerSettings
			{
				ContractResolver = new CamelCasePropertyNamesContractResolver()
			};
			var emailJson = new StringContent(JsonConvert.SerializeObject(email, jsonSerializerSettings), Encoding.UTF8, "application/json");

			String url = "https://portal.precast.org/api/v1/outlook/archived-emails/"
					+ entityId + "/" + entryId + "?apiToken=MUg@R*A8jgtwY$aQXv3J";
			try 
			{
				var client = new HttpClient();
				var request = new HttpRequestMessage(HttpMethod.Put, url);
				request.Content = emailJson;
				var response = client.SendAsync(request).Result;
				entityIdProperty.Value += ", " + entityId;
				mailItem.MessageClass = "IPM.Note.Phoenix";
				mailItem.Save();
			}
			catch
			{
				System.Windows.Forms.MessageBox.Show("Failed to archive the email. Please try again.");
			}
		}
    }
}