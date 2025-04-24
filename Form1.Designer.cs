
namespace PullAMSData3
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
            this.btnGetAMSData = new System.Windows.Forms.Button();
            this.btnODBC = new System.Windows.Forms.Button();
            this.dataGridViewAMSTables = new System.Windows.Forms.DataGridView();
            this.chkboxCountsOnly = new System.Windows.Forms.CheckBox();
            this.btnSelectAllRows = new System.Windows.Forms.Button();
            this.txtExplain = new System.Windows.Forms.TextBox();
            this.btnShowAMS = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewAMSTables)).BeginInit();
            this.SuspendLayout();
            // 
            // btnGetAMSData
            // 
            this.btnGetAMSData.Location = new System.Drawing.Point(38, 73);
            this.btnGetAMSData.Name = "btnGetAMSData";
            this.btnGetAMSData.Size = new System.Drawing.Size(122, 78);
            this.btnGetAMSData.TabIndex = 0;
            this.btnGetAMSData.Text = "Get AMS (OLEDB) Data";
            this.btnGetAMSData.UseVisualStyleBackColor = true;
            this.btnGetAMSData.Click += new System.EventHandler(this.btnGetAMSData_Click);
            // 
            // btnODBC
            // 
            this.btnODBC.Enabled = false;
            this.btnODBC.Location = new System.Drawing.Point(38, 434);
            this.btnODBC.Name = "btnODBC";
            this.btnODBC.Size = new System.Drawing.Size(176, 61);
            this.btnODBC.TabIndex = 1;
            this.btnODBC.Text = "Show Catalog Master Tables";
            this.btnODBC.UseVisualStyleBackColor = true;
            this.btnODBC.Visible = false;
            this.btnODBC.Click += new System.EventHandler(this.btnODBC_Click);
            // 
            // dataGridViewAMSTables
            // 
            this.dataGridViewAMSTables.AllowUserToAddRows = false;
            this.dataGridViewAMSTables.AllowUserToDeleteRows = false;
            this.dataGridViewAMSTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewAMSTables.Location = new System.Drawing.Point(316, 73);
            this.dataGridViewAMSTables.Name = "dataGridViewAMSTables";
            this.dataGridViewAMSTables.RowHeadersWidth = 62;
            this.dataGridViewAMSTables.RowTemplate.Height = 28;
            this.dataGridViewAMSTables.Size = new System.Drawing.Size(1008, 547);
            this.dataGridViewAMSTables.TabIndex = 2;
            // 
            // chkboxCountsOnly
            // 
            this.chkboxCountsOnly.AutoSize = true;
            this.chkboxCountsOnly.Location = new System.Drawing.Point(164, 101);
            this.chkboxCountsOnly.Name = "chkboxCountsOnly";
            this.chkboxCountsOnly.Size = new System.Drawing.Size(146, 24);
            this.chkboxCountsOnly.TabIndex = 3;
            this.chkboxCountsOnly.Text = "Get counts only";
            this.chkboxCountsOnly.UseVisualStyleBackColor = true;
            // 
            // btnSelectAllRows
            // 
            this.btnSelectAllRows.BackColor = System.Drawing.SystemColors.Highlight;
            this.btnSelectAllRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectAllRows.Location = new System.Drawing.Point(310, 13);
            this.btnSelectAllRows.Name = "btnSelectAllRows";
            this.btnSelectAllRows.Size = new System.Drawing.Size(149, 54);
            this.btnSelectAllRows.TabIndex = 4;
            this.btnSelectAllRows.Text = "Select All Rows";
            this.btnSelectAllRows.UseVisualStyleBackColor = false;
            this.btnSelectAllRows.Click += new System.EventHandler(this.btnSelectAllRows_Click);
            // 
            // txtExplain
            // 
            this.txtExplain.Location = new System.Drawing.Point(27, 175);
            this.txtExplain.Multiline = true;
            this.txtExplain.Name = "txtExplain";
            this.txtExplain.ReadOnly = true;
            this.txtExplain.Size = new System.Drawing.Size(232, 180);
            this.txtExplain.TabIndex = 5;
            this.txtExplain.Text = resources.GetString("txtExplain.Text");
            // 
            // btnShowAMS
            // 
            this.btnShowAMS.Location = new System.Drawing.Point(38, 13);
            this.btnShowAMS.Name = "btnShowAMS";
            this.btnShowAMS.Size = new System.Drawing.Size(126, 54);
            this.btnShowAMS.TabIndex = 6;
            this.btnShowAMS.Text = "Show AMS Tables";
            this.btnShowAMS.UseVisualStyleBackColor = true;
            this.btnShowAMS.Click += new System.EventHandler(this.btnShowAMS_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(751, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 32);
            this.label1.TabIndex = 7;
            this.label1.Text = "AMS";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(890, 19);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(162, 42);
            this.button1.TabIndex = 8;
            this.button1.Text = "Test Workaround";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(915, -4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 20);
            this.label2.TabIndex = 9;
            this.label2.Text = "This works!!";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1343, 603);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnShowAMS);
            this.Controls.Add(this.txtExplain);
            this.Controls.Add(this.btnSelectAllRows);
            this.Controls.Add(this.chkboxCountsOnly);
            this.Controls.Add(this.dataGridViewAMSTables);
            this.Controls.Add(this.btnODBC);
            this.Controls.Add(this.btnGetAMSData);
            this.Name = "Form1";
            this.Text = "Get AMS Data";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewAMSTables)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGetAMSData;
        private System.Windows.Forms.Button btnODBC;
        private System.Windows.Forms.DataGridView dataGridViewAMSTables;
        private System.Windows.Forms.CheckBox chkboxCountsOnly;
        private System.Windows.Forms.Button btnSelectAllRows;
        private System.Windows.Forms.TextBox txtExplain;
        private System.Windows.Forms.Button btnShowAMS;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
    }
}

