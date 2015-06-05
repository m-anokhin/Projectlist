namespace ExcelProject
{
    partial class ManualFormForGiBProjects
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.ExcelGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ExcelGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.MinimumSize = new System.Drawing.Size(1400, 50);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.ExcelGridView);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.AutoScroll = true;
            this.splitContainer1.Size = new System.Drawing.Size(1400, 561);
            this.splitContainer1.SplitterDistance = 386;
            this.splitContainer1.TabIndex = 0;
            // 
            // ExcelGridView
            // 
            this.ExcelGridView.AllowUserToAddRows = false;
            this.ExcelGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.ExcelGridView.BackgroundColor = System.Drawing.SystemColors.Control;
            this.ExcelGridView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.ExcelGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ExcelGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ExcelGridView.Location = new System.Drawing.Point(0, 0);
            this.ExcelGridView.Margin = new System.Windows.Forms.Padding(0);
            this.ExcelGridView.MultiSelect = false;
            this.ExcelGridView.Name = "ExcelGridView";
            this.ExcelGridView.ReadOnly = true;
            this.ExcelGridView.RowHeadersVisible = false;
            this.ExcelGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ExcelGridView.Size = new System.Drawing.Size(1400, 386);
            this.ExcelGridView.TabIndex = 0;
            this.ExcelGridView.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.ExcelGridView_DataBindingComplete);
            this.ExcelGridView.SelectionChanged += new System.EventHandler(this.ExcelGridView_SelectionChanged);
            // 
            // ManualFormForGiBProjects
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1334, 561);
            this.Controls.Add(this.splitContainer1);
            this.MaximumSize = new System.Drawing.Size(1800, 600);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(1300, 550);
            this.Name = "ManualFormForGiBProjects";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ManualFormForGiBProjects";
            this.Load += new System.EventHandler(this.ManualFormForGiBProjects_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ExcelGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView ExcelGridView;
    }
}