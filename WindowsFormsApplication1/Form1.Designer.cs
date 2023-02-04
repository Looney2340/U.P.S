namespace WindowsFormsApplication1
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
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.Choosedriver = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.paidDay = new System.Windows.Forms.TextBox();
            this.planDay = new System.Windows.Forms.TextBox();
            this.ovUn = new System.Windows.Forms.TextBox();
            this.amTime = new System.Windows.Forms.TextBox();
            this.pmTime = new System.Windows.Forms.TextBox();
            this.sPORH = new System.Windows.Forms.TextBox();
            this.delStops = new System.Windows.Forms.TextBox();
            this.delPkgs = new System.Windows.Forms.TextBox();
            this.puStops = new System.Windows.Forms.TextBox();
            this.puPkgs = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dteTime = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.sLIC = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.bLDG = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.employeeID = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.BackColor = System.Drawing.SystemColors.Window;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(48, 154);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboBox1.MaxDropDownItems = 25;
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(220, 28);
            this.comboBox1.Sorted = true;
            this.comboBox1.TabIndex = 0;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // Choosedriver
            // 
            this.Choosedriver.AutoSize = true;
            this.Choosedriver.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Choosedriver.Location = new System.Drawing.Point(48, 120);
            this.Choosedriver.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.Choosedriver.Name = "Choosedriver";
            this.Choosedriver.Size = new System.Drawing.Size(224, 29);
            this.Choosedriver.TabIndex = 1;
            this.Choosedriver.Text = "CHOOSE DRIVER";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(555, 56);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(251, 29);
            this.label2.TabIndex = 2;
            this.label2.Text = "2 WEEK AVERAGES";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Gold;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1214, 1);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(47, 42);
            this.button1.TabIndex = 3;
            this.button1.Text = "X";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(315, 122);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 29);
            this.label1.TabIndex = 4;
            this.label1.Text = "PaidDay\r\n";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(443, 122);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 29);
            this.label3.TabIndex = 5;
            this.label3.Text = "PlanDay";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(584, 122);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(163, 29);
            this.label4.TabIndex = 6;
            this.label4.Text = "OverAllowed";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(767, 122);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(117, 29);
            this.label5.TabIndex = 7;
            this.label5.Text = "AM Time";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(909, 122);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(118, 29);
            this.label6.TabIndex = 8;
            this.label6.Text = "PM Time";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1071, 122);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(103, 29);
            this.label7.TabIndex = 9;
            this.label7.Text = "SPORH";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(298, 316);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(127, 29);
            this.label8.TabIndex = 10;
            this.label8.Text = "Del Stops";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(457, 316);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(118, 29);
            this.label9.TabIndex = 11;
            this.label9.Text = "Del Pkgs";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(810, 316);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(116, 29);
            this.label11.TabIndex = 13;
            this.label11.Text = "Pick Ups";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(959, 316);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(153, 29);
            this.label12.TabIndex = 14;
            this.label12.Text = "Pick Up Pcs";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(462, 54);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(86, 31);
            this.button2.TabIndex = 15;
            this.button2.Text = "Pull PFT";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(47, 88);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(221, 28);
            this.comboBox2.TabIndex = 16;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(48, 56);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(220, 29);
            this.label13.TabIndex = 17;
            this.label13.Text = "CHOOSE ROUTE";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(370, 54);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(86, 31);
            this.button3.TabIndex = 18;
            this.button3.Text = "Update";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // paidDay
            // 
            this.paidDay.Location = new System.Drawing.Point(320, 156);
            this.paidDay.Name = "paidDay";
            this.paidDay.Size = new System.Drawing.Size(100, 26);
            this.paidDay.TabIndex = 19;
            // 
            // planDay
            // 
            this.planDay.Location = new System.Drawing.Point(448, 156);
            this.planDay.Name = "planDay";
            this.planDay.Size = new System.Drawing.Size(100, 26);
            this.planDay.TabIndex = 20;
            // 
            // ovUn
            // 
            this.ovUn.Location = new System.Drawing.Point(619, 156);
            this.ovUn.Name = "ovUn";
            this.ovUn.Size = new System.Drawing.Size(100, 26);
            this.ovUn.TabIndex = 21;
            // 
            // amTime
            // 
            this.amTime.Location = new System.Drawing.Point(772, 156);
            this.amTime.Name = "amTime";
            this.amTime.Size = new System.Drawing.Size(100, 26);
            this.amTime.TabIndex = 22;
            // 
            // pmTime
            // 
            this.pmTime.Location = new System.Drawing.Point(914, 156);
            this.pmTime.Name = "pmTime";
            this.pmTime.Size = new System.Drawing.Size(100, 26);
            this.pmTime.TabIndex = 23;
            // 
            // sPORH
            // 
            this.sPORH.Location = new System.Drawing.Point(1074, 156);
            this.sPORH.Name = "sPORH";
            this.sPORH.Size = new System.Drawing.Size(100, 26);
            this.sPORH.TabIndex = 24;
            // 
            // delStops
            // 
            this.delStops.Location = new System.Drawing.Point(320, 348);
            this.delStops.Name = "delStops";
            this.delStops.Size = new System.Drawing.Size(100, 26);
            this.delStops.TabIndex = 25;
            // 
            // delPkgs
            // 
            this.delPkgs.Location = new System.Drawing.Point(462, 348);
            this.delPkgs.Name = "delPkgs";
            this.delPkgs.Size = new System.Drawing.Size(100, 26);
            this.delPkgs.TabIndex = 26;
            // 
            // puStops
            // 
            this.puStops.Location = new System.Drawing.Point(815, 348);
            this.puStops.Name = "puStops";
            this.puStops.Size = new System.Drawing.Size(100, 26);
            this.puStops.TabIndex = 27;
            // 
            // puPkgs
            // 
            this.puPkgs.Location = new System.Drawing.Point(986, 348);
            this.puPkgs.Name = "puPkgs";
            this.puPkgs.Size = new System.Drawing.Size(100, 26);
            this.puPkgs.TabIndex = 28;
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.Gold;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(1159, 1);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(47, 42);
            this.button4.TabIndex = 29;
            this.button4.Text = "_";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dteTime);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.sLIC);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.bLDG);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Location = new System.Drawing.Point(12, 657);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(669, 71);
            this.groupBox1.TabIndex = 30;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = " ..:: CONFiG ::.. ";
            // 
            // dteTime
            // 
            this.dteTime.Location = new System.Drawing.Point(447, 28);
            this.dteTime.Name = "dteTime";
            this.dteTime.Size = new System.Drawing.Size(213, 26);
            this.dteTime.TabIndex = 5;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(354, 31);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(87, 20);
            this.label16.TabIndex = 4;
            this.label16.Text = "Data Date:";
            // 
            // sLIC
            // 
            this.sLIC.Location = new System.Drawing.Point(223, 28);
            this.sLIC.Name = "sLIC";
            this.sLIC.Size = new System.Drawing.Size(100, 26);
            this.sLIC.TabIndex = 3;
            this.sLIC.Text = "1166";
            this.sLIC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(172, 31);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(45, 20);
            this.label15.TabIndex = 2;
            this.label15.Text = "SLIC";
            // 
            // bLDG
            // 
            this.bLDG.Location = new System.Drawing.Point(66, 28);
            this.bLDG.Name = "bLDG";
            this.bLDG.Size = new System.Drawing.Size(100, 26);
            this.bLDG.TabIndex = 1;
            this.bLDG.Text = "NYSPR";
            this.bLDG.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.bLDG.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 31);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(54, 20);
            this.label14.TabIndex = 0;
            this.label14.Text = "BLDG";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(614, 316);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(162, 29);
            this.label10.TabIndex = 12;
            this.label10.Text = "Employee ID";
            // 
            // employeeID
            // 
            this.employeeID.Location = new System.Drawing.Point(619, 348);
            this.employeeID.Name = "employeeID";
            this.employeeID.Size = new System.Drawing.Size(157, 26);
            this.employeeID.TabIndex = 31;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Beige;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(1264, 740);
            this.Controls.Add(this.employeeID);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.puPkgs);
            this.Controls.Add(this.puStops);
            this.Controls.Add(this.delPkgs);
            this.Controls.Add(this.delStops);
            this.Controls.Add(this.sPORH);
            this.Controls.Add(this.pmTime);
            this.Controls.Add(this.amTime);
            this.Controls.Add(this.ovUn);
            this.Controls.Add(this.planDay);
            this.Controls.Add(this.paidDay);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Choosedriver);
            this.Controls.Add(this.comboBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PFT 2 Week Trend";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label Choosedriver;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox paidDay;
        private System.Windows.Forms.TextBox planDay;
        private System.Windows.Forms.TextBox ovUn;
        private System.Windows.Forms.TextBox amTime;
        private System.Windows.Forms.TextBox pmTime;
        private System.Windows.Forms.TextBox sPORH;
        private System.Windows.Forms.TextBox delStops;
        private System.Windows.Forms.TextBox delPkgs;
        private System.Windows.Forms.TextBox puStops;
        private System.Windows.Forms.TextBox puPkgs;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox bLDG;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox sLIC;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox dteTime;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox employeeID;
    }
}

