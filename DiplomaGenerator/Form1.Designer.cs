namespace DiplomaGenerator
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.comboBoxCourseModule = new System.Windows.Forms.ComboBox();
            this.txtNume = new System.Windows.Forms.TextBox();
            this.txtCNP = new System.Windows.Forms.TextBox();
            this.dateTimePickerEliberare = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerExpirare = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.openFileDialogExcel = new System.Windows.Forms.OpenFileDialog();
            this.txtExcelFilePath = new System.Windows.Forms.TextBox();
            this.btnChooseExcel = new System.Windows.Forms.Button();
            this.progressBarGenerating = new System.Windows.Forms.ProgressBar();
            this.label9 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBoxCourseModule
            // 
            this.comboBoxCourseModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCourseModule.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.comboBoxCourseModule.FormattingEnabled = true;
            this.comboBoxCourseModule.Location = new System.Drawing.Point(16, 301);
            this.comboBoxCourseModule.Name = "comboBoxCourseModule";
            this.comboBoxCourseModule.Size = new System.Drawing.Size(156, 28);
            this.comboBoxCourseModule.TabIndex = 0;
            this.comboBoxCourseModule.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // txtNume
            // 
            this.txtNume.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.txtNume.Location = new System.Drawing.Point(16, 356);
            this.txtNume.Name = "txtNume";
            this.txtNume.Size = new System.Drawing.Size(288, 27);
            this.txtNume.TabIndex = 1;
            // 
            // txtCNP
            // 
            this.txtCNP.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.txtCNP.Location = new System.Drawing.Point(16, 406);
            this.txtCNP.Name = "txtCNP";
            this.txtCNP.Size = new System.Drawing.Size(288, 27);
            this.txtCNP.TabIndex = 3;
            // 
            // dateTimePickerEliberare
            // 
            this.dateTimePickerEliberare.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickerEliberare.Location = new System.Drawing.Point(16, 457);
            this.dateTimePickerEliberare.Name = "dateTimePickerEliberare";
            this.dateTimePickerEliberare.Size = new System.Drawing.Size(200, 22);
            this.dateTimePickerEliberare.TabIndex = 6;
            this.dateTimePickerEliberare.ValueChanged += new System.EventHandler(this.dateTimePickerEliberare_ValueChanged);
            // 
            // dateTimePickerExpirare
            // 
            this.dateTimePickerExpirare.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickerExpirare.Location = new System.Drawing.Point(16, 508);
            this.dateTimePickerExpirare.Name = "dateTimePickerExpirare";
            this.dateTimePickerExpirare.Size = new System.Drawing.Size(200, 22);
            this.dateTimePickerExpirare.TabIndex = 7;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(12, 540);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(426, 50);
            this.button1.TabIndex = 8;
            this.button1.Text = "GENEREAZA DIPLOME";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(16, 141);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 20);
            this.label1.TabIndex = 10;
            this.label1.Text = "Alege fisierul Excel";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(16, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(138, 21);
            this.label2.TabIndex = 11;
            this.label2.Text = "Varianta 1 - Auto";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(15, 205);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 30);
            this.label3.TabIndex = 12;
            this.label3.Text = "SAU";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.White;
            this.label10.Location = new System.Drawing.Point(16, 248);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(159, 21);
            this.label10.TabIndex = 19;
            this.label10.Text = "Varianta 2 - Manual";
            this.label10.Click += new System.EventHandler(this.label10_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(16, 279);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(157, 20);
            this.label4.TabIndex = 20;
            this.label4.Text = "Alege model Diploma";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(16, 332);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(102, 20);
            this.label5.TabIndex = 21;
            this.label5.Text = "Nume Cursant";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(16, 383);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 20);
            this.label6.TabIndex = 22;
            this.label6.Text = "CNP/ID";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(16, 435);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(96, 20);
            this.label7.TabIndex = 23;
            this.label7.Text = "Data Emitere";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(16, 482);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(99, 20);
            this.label8.TabIndex = 24;
            this.label8.Text = "Data Expirare";
            // 
            // txtExcelFilePath
            // 
            this.txtExcelFilePath.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExcelFilePath.Location = new System.Drawing.Point(152, 166);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.ReadOnly = true;
            this.txtExcelFilePath.Size = new System.Drawing.Size(286, 22);
            this.txtExcelFilePath.TabIndex = 26;
            // 
            // btnChooseExcel
            // 
            this.btnChooseExcel.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChooseExcel.Location = new System.Drawing.Point(16, 164);
            this.btnChooseExcel.Name = "btnChooseExcel";
            this.btnChooseExcel.Size = new System.Drawing.Size(130, 25);
            this.btnChooseExcel.TabIndex = 27;
            this.btnChooseExcel.Text = "Alege...";
            this.btnChooseExcel.UseVisualStyleBackColor = true;
            this.btnChooseExcel.Click += new System.EventHandler(this.btnChooseExcel_Click);
            // 
            // progressBarGenerating
            // 
            this.progressBarGenerating.Location = new System.Drawing.Point(12, 540);
            this.progressBarGenerating.Name = "progressBarGenerating";
            this.progressBarGenerating.Size = new System.Drawing.Size(430, 50);
            this.progressBarGenerating.TabIndex = 28;
            this.progressBarGenerating.Visible = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.BackColor = System.Drawing.Color.Transparent;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 22F, System.Drawing.FontStyle.Bold);
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(13, 12);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(287, 41);
            this.label9.TabIndex = 29;
            this.label9.Text = "Generator Diplome";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.BackColor = System.Drawing.Color.Transparent;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.label11.ForeColor = System.Drawing.Color.White;
            this.label11.Location = new System.Drawing.Point(17, 53);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(281, 26);
            this.label11.TabIndex = 30;
            this.label11.Text = "A se folosi cu foaia de calcul Generator Diplome.xlsx \r\ngithub.com/theodor-ubp\t\r\n";
            this.label11.Click += new System.EventHandler(this.label11_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.button2.Location = new System.Drawing.Point(12, 607);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(134, 42);
            this.button2.TabIndex = 31;
            this.button2.Text = "Citește-mă!";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.button3.Location = new System.Drawing.Point(155, 607);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(146, 42);
            this.button3.TabIndex = 32;
            this.button3.Text = "Cum creez sabloane?";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.button4.Location = new System.Drawing.Point(310, 607);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(132, 42);
            this.button4.TabIndex = 33;
            this.button4.Text = "Cum creez xlsx pentru Auto?";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(450, 661);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.progressBarGenerating);
            this.Controls.Add(this.btnChooseExcel);
            this.Controls.Add(this.txtExcelFilePath);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dateTimePickerExpirare);
            this.Controls.Add(this.dateTimePickerEliberare);
            this.Controls.Add(this.txtCNP);
            this.Controls.Add(this.txtNume);
            this.Controls.Add(this.comboBoxCourseModule);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(466, 700);
            this.MinimumSize = new System.Drawing.Size(466, 700);
            this.Name = "Form1";
            this.Text = "Generator Diplome v1.10";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxCourseModule;
        private System.Windows.Forms.TextBox txtNume;
        private System.Windows.Forms.TextBox txtCNP;
        private System.Windows.Forms.DateTimePicker dateTimePickerEliberare;
        private System.Windows.Forms.DateTimePicker dateTimePickerExpirare;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.OpenFileDialog openFileDialogExcel;
        private System.Windows.Forms.TextBox txtExcelFilePath;
        private System.Windows.Forms.Button btnChooseExcel;
        private System.Windows.Forms.ProgressBar progressBarGenerating;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
    }
}

