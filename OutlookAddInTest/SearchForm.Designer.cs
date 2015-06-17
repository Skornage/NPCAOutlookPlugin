namespace OutlookAddInTest
{
    partial class SearchForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.infoInput = new System.Windows.Forms.TextBox();
            this.emailInput = new System.Windows.Forms.TextBox();
            this.nameInput = new System.Windows.Forms.TextBox();
            this.typeInput = new System.Windows.Forms.TextBox();
            this.idInput = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.infoInput);
            this.groupBox1.Controls.Add(this.emailInput);
            this.groupBox1.Controls.Add(this.nameInput);
            this.groupBox1.Controls.Add(this.typeInput);
            this.groupBox1.Controls.Add(this.idInput);
            this.groupBox1.Location = new System.Drawing.Point(13, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(325, 273);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Fill out any number of fields";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(27, 214);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 17);
            this.label5.TabIndex = 9;
            this.label5.Text = "info";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 172);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 17);
            this.label4.TabIndex = 8;
            this.label4.Text = "email";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 129);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "type";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "id";
            // 
            // infoInput
            // 
            this.infoInput.Location = new System.Drawing.Point(141, 214);
            this.infoInput.Name = "infoInput";
            this.infoInput.Size = new System.Drawing.Size(160, 22);
            this.infoInput.TabIndex = 4;
            // 
            // emailInput
            // 
            this.emailInput.Location = new System.Drawing.Point(141, 172);
            this.emailInput.Name = "emailInput";
            this.emailInput.Size = new System.Drawing.Size(160, 22);
            this.emailInput.TabIndex = 3;
            // 
            // nameInput
            // 
            this.nameInput.Location = new System.Drawing.Point(141, 129);
            this.nameInput.Name = "nameInput";
            this.nameInput.Size = new System.Drawing.Size(160, 22);
            this.nameInput.TabIndex = 2;
            // 
            // typeInput
            // 
            this.typeInput.Location = new System.Drawing.Point(141, 82);
            this.typeInput.Name = "typeInput";
            this.typeInput.Size = new System.Drawing.Size(160, 22);
            this.typeInput.TabIndex = 1;
            // 
            // idInput
            // 
            this.idInput.Location = new System.Drawing.Point(141, 40);
            this.idInput.Name = "idInput";
            this.idInput.Size = new System.Drawing.Size(160, 22);
            this.idInput.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(154, 316);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(87, 29);
            this.button1.TabIndex = 1;
            this.button1.Text = "Search";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(251, 316);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(87, 29);
            this.button2.TabIndex = 2;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // SearchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(366, 371);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Name = "SearchForm";
            this.Text = "Advanced Search";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox infoInput;
        private System.Windows.Forms.TextBox emailInput;
        private System.Windows.Forms.TextBox nameInput;
        private System.Windows.Forms.TextBox typeInput;
        private System.Windows.Forms.TextBox idInput;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}