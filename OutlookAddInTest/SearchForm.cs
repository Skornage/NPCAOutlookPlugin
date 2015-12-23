using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddInTest
{
    
    public partial class SearchForm : Form
    {
        MainForm form;
        public SearchForm(MainForm form)
        {
            this.form = form;
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string id = idInput.Text.Trim();
            string type = typeInput.Text.Trim();
            string name = nameInput.Text.Trim();
            string email = emailInput.Text.Trim();
            this.form.advancedSearch(id, type, name, email);
            this.Close();
        }
    }
}
