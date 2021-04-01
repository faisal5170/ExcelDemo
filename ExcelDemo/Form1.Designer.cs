namespace ExcelDemo
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
            this.btnReadExcel = new System.Windows.Forms.Button();
            this.openExcel = new System.Windows.Forms.Button();
            this.insertExcel = new System.Windows.Forms.Button();
            this.clearExcel = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnReadExcel
            // 
            this.btnReadExcel.Location = new System.Drawing.Point(130, 35);
            this.btnReadExcel.Name = "btnReadExcel";
            this.btnReadExcel.Size = new System.Drawing.Size(75, 23);
            this.btnReadExcel.TabIndex = 0;
            this.btnReadExcel.Text = "Read Excel";
            this.btnReadExcel.UseVisualStyleBackColor = true;
            this.btnReadExcel.Click += new System.EventHandler(this.btnReadExcel_Click);
            // 
            // openExcel
            // 
            this.openExcel.Location = new System.Drawing.Point(270, 34);
            this.openExcel.Name = "openExcel";
            this.openExcel.Size = new System.Drawing.Size(99, 23);
            this.openExcel.TabIndex = 1;
            this.openExcel.Text = "Choose Excel";
            this.openExcel.UseVisualStyleBackColor = true;
            this.openExcel.Click += new System.EventHandler(this.openExcel_Click);
            // 
            // insertExcel
            // 
            this.insertExcel.Location = new System.Drawing.Point(583, 35);
            this.insertExcel.Name = "insertExcel";
            this.insertExcel.Size = new System.Drawing.Size(135, 23);
            this.insertExcel.TabIndex = 2;
            this.insertExcel.Text = "Confirm Insert Data";
            this.insertExcel.UseVisualStyleBackColor = true;
            this.insertExcel.Click += new System.EventHandler(this.insertExcel_Click);
            // 
            // clearExcel
            // 
            this.clearExcel.Location = new System.Drawing.Point(399, 34);
            this.clearExcel.Name = "clearExcel";
            this.clearExcel.Size = new System.Drawing.Size(145, 23);
            this.clearExcel.TabIndex = 3;
            this.clearExcel.Text = "Clear Current Data";
            this.clearExcel.UseVisualStyleBackColor = true;
            this.clearExcel.Click += new System.EventHandler(this.clearExcel_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(13, 85);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(775, 353);
            this.dataGridView1.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.clearExcel);
            this.Controls.Add(this.insertExcel);
            this.Controls.Add(this.openExcel);
            this.Controls.Add(this.btnReadExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnReadExcel;
        private System.Windows.Forms.Button openExcel;
        private System.Windows.Forms.Button insertExcel;
        private System.Windows.Forms.Button clearExcel;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

