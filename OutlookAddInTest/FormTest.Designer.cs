﻿namespace OutlookAddInTest
{
	partial class FormTest
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.Cancel = new System.Windows.Forms.Button();
			this.Ok = new System.Windows.Forms.Button();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.button3 = new System.Windows.Forms.Button();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.listBox2 = new System.Windows.Forms.ListBox();
			this.SuspendLayout();
			// 
			// listBox1
			// 
			this.listBox1.FormattingEnabled = true;
			this.listBox1.ItemHeight = 16;
			this.listBox1.Location = new System.Drawing.Point(12, 123);
			this.listBox1.Name = "listBox1";
			this.listBox1.Size = new System.Drawing.Size(618, 260);
			this.listBox1.TabIndex = 0;
			// 
			// Cancel
			// 
			this.Cancel.Location = new System.Drawing.Point(540, 399);
			this.Cancel.Name = "Cancel";
			this.Cancel.Size = new System.Drawing.Size(90, 29);
			this.Cancel.TabIndex = 1;
			this.Cancel.Text = "Cancel";
			this.Cancel.UseVisualStyleBackColor = true;
			this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
			// 
			// Ok
			// 
			this.Ok.Location = new System.Drawing.Point(436, 399);
			this.Ok.Name = "Ok";
			this.Ok.Size = new System.Drawing.Size(87, 29);
			this.Ok.TabIndex = 2;
			this.Ok.Text = "Ok";
			this.Ok.UseVisualStyleBackColor = true;
			this.Ok.Click += new System.EventHandler(this.Ok_Click);
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(402, 34);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(132, 22);
			this.textBox1.TabIndex = 5;
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(540, 34);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(62, 22);
			this.button3.TabIndex = 6;
			this.button3.Text = "Search";
			this.button3.UseVisualStyleBackColor = true;
			// 
			// comboBox1
			// 
			this.comboBox1.DisplayMember = "Test";
			this.comboBox1.FormattingEnabled = true;
			this.comboBox1.Items.AddRange(new object[] {
            "Account",
            "Contact"});
			this.comboBox1.Location = new System.Drawing.Point(12, 32);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(278, 24);
			this.comboBox1.TabIndex = 7;
			this.comboBox1.ValueMember = "Test";
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
			// 
			// listBox2
			// 
			this.listBox2.FormattingEnabled = true;
			this.listBox2.ItemHeight = 16;
			this.listBox2.Location = new System.Drawing.Point(12, 97);
			this.listBox2.Name = "listBox2";
			this.listBox2.Size = new System.Drawing.Size(618, 20);
			this.listBox2.TabIndex = 8;
			this.listBox2.SelectedIndexChanged += new System.EventHandler(this.listBox2_SelectedIndexChanged);
			// 
			// FormTest
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(642, 438);
			this.Controls.Add(this.listBox2);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.textBox1);
			this.Controls.Add(this.Ok);
			this.Controls.Add(this.Cancel);
			this.Controls.Add(this.listBox1);
			this.Name = "FormTest";
			this.Text = "FormTest";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.Button Cancel;
		private System.Windows.Forms.Button Ok;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.ListBox listBox2;
	}
}