namespace ExcelProject
{
    partial class LogForm
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
            this.LogDataGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.LogDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // LogDataGridView
            // 
            this.LogDataGridView.AllowUserToAddRows = false;
            this.LogDataGridView.AllowUserToDeleteRows = false;
            this.LogDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.LogDataGridView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.LogDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LogDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LogDataGridView.GridColor = System.Drawing.SystemColors.Window;
            this.LogDataGridView.Location = new System.Drawing.Point(0, 0);
            this.LogDataGridView.MultiSelect = false;
            this.LogDataGridView.Name = "LogDataGridView";
            this.LogDataGridView.ReadOnly = true;
            this.LogDataGridView.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.LogDataGridView.RowHeadersVisible = false;
            this.LogDataGridView.Size = new System.Drawing.Size(644, 261);
            this.LogDataGridView.TabIndex = 0;
            this.LogDataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // LogForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(644, 261);
            this.Controls.Add(this.LogDataGridView);
            this.MaximumSize = new System.Drawing.Size(660, 500);
            this.MinimumSize = new System.Drawing.Size(660, 0);
            this.Name = "LogForm";
            this.Text = "LogForm";
            this.Load += new System.EventHandler(this.LogForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.LogDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView LogDataGridView;
    }
}