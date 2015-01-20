namespace VisualTexcel
{
    partial class GUI
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
            this.rangeStart = new System.Windows.Forms.Button();
            this.listStart = new System.Windows.Forms.Button();
            this.singleStart = new System.Windows.Forms.Button();
            this.RangeofFilesLabel = new System.Windows.Forms.Label();
            this.minFileNum = new System.Windows.Forms.NumericUpDown();
            this.minFileNumLabel = new System.Windows.Forms.Label();
            this.maxFileNumLabel = new System.Windows.Forms.Label();
            this.maxFileNum = new System.Windows.Forms.NumericUpDown();
            this.ListofFilesLabel = new System.Windows.Forms.Label();
            this.listInstructionLabel = new System.Windows.Forms.Label();
            this.SingleFileLabel = new System.Windows.Forms.Label();
            this.singleFileNum = new System.Windows.Forms.NumericUpDown();
            this.singleFileSelectLabel = new System.Windows.Forms.Label();
            this.listString = new System.Windows.Forms.TextBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.ProgressLabel = new System.Windows.Forms.Label();
            this.progressMessage1 = new System.Windows.Forms.Label();
            this.progressMessage2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menu = new System.Windows.Forms.ToolStripMenuItem();
            this.menu_help = new System.Windows.Forms.ToolStripMenuItem();
            this.menu_about = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.minFileNum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.maxFileNum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.singleFileNum)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rangeStart
            // 
            this.rangeStart.AutoSize = true;
            this.rangeStart.Location = new System.Drawing.Point(17, 137);
            this.rangeStart.Name = "rangeStart";
            this.rangeStart.Size = new System.Drawing.Size(213, 24);
            this.rangeStart.TabIndex = 0;
            this.rangeStart.Text = "Process Range";
            this.rangeStart.UseVisualStyleBackColor = true;
            this.rangeStart.Click += new System.EventHandler(this.rangeStart_Click);
            // 
            // listStart
            // 
            this.listStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.listStart.AutoSize = true;
            this.listStart.Location = new System.Drawing.Point(18, 257);
            this.listStart.Name = "listStart";
            this.listStart.Size = new System.Drawing.Size(213, 24);
            this.listStart.TabIndex = 1;
            this.listStart.Text = "Process List";
            this.listStart.UseVisualStyleBackColor = true;
            this.listStart.Click += new System.EventHandler(this.listStart_Click);
            // 
            // singleStart
            // 
            this.singleStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.singleStart.AutoSize = true;
            this.singleStart.Location = new System.Drawing.Point(279, 137);
            this.singleStart.Name = "singleStart";
            this.singleStart.Size = new System.Drawing.Size(213, 24);
            this.singleStart.TabIndex = 2;
            this.singleStart.Text = "Process File";
            this.singleStart.UseVisualStyleBackColor = true;
            this.singleStart.Click += new System.EventHandler(this.singleStart_Click);
            // 
            // RangeofFilesLabel
            // 
            this.RangeofFilesLabel.AutoSize = true;
            this.RangeofFilesLabel.Font = new System.Drawing.Font("Calibri", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RangeofFilesLabel.Location = new System.Drawing.Point(11, 38);
            this.RangeofFilesLabel.Name = "RangeofFilesLabel";
            this.RangeofFilesLabel.Size = new System.Drawing.Size(165, 33);
            this.RangeofFilesLabel.TabIndex = 3;
            this.RangeofFilesLabel.Text = "Range of Files";
            // 
            // minFileNum
            // 
            this.minFileNum.AutoSize = true;
            this.minFileNum.Location = new System.Drawing.Point(110, 74);
            this.minFileNum.Name = "minFileNum";
            this.minFileNum.Size = new System.Drawing.Size(120, 22);
            this.minFileNum.TabIndex = 4;
            this.minFileNum.ValueChanged += new System.EventHandler(this.minFileNum_ValueChanged);
            // 
            // minFileNumLabel
            // 
            this.minFileNumLabel.AutoSize = true;
            this.minFileNumLabel.Location = new System.Drawing.Point(14, 76);
            this.minFileNumLabel.Name = "minFileNumLabel";
            this.minFileNumLabel.Size = new System.Drawing.Size(90, 14);
            this.minFileNumLabel.TabIndex = 5;
            this.minFileNumLabel.Text = "File to start at: ";
            // 
            // maxFileNumLabel
            // 
            this.maxFileNumLabel.AutoSize = true;
            this.maxFileNumLabel.Location = new System.Drawing.Point(15, 108);
            this.maxFileNumLabel.Name = "maxFileNumLabel";
            this.maxFileNumLabel.Size = new System.Drawing.Size(89, 14);
            this.maxFileNumLabel.TabIndex = 6;
            this.maxFileNumLabel.Text = "File to end at:  ";
            // 
            // maxFileNum
            // 
            this.maxFileNum.AutoSize = true;
            this.maxFileNum.Location = new System.Drawing.Point(110, 106);
            this.maxFileNum.Name = "maxFileNum";
            this.maxFileNum.Size = new System.Drawing.Size(120, 22);
            this.maxFileNum.TabIndex = 7;
            this.maxFileNum.ValueChanged += new System.EventHandler(this.maxFileNum_ValueChanged);
            // 
            // ListofFilesLabel
            // 
            this.ListofFilesLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.ListofFilesLabel.AutoSize = true;
            this.ListofFilesLabel.Font = new System.Drawing.Font("Calibri", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ListofFilesLabel.Location = new System.Drawing.Point(12, 170);
            this.ListofFilesLabel.Name = "ListofFilesLabel";
            this.ListofFilesLabel.Size = new System.Drawing.Size(134, 33);
            this.ListofFilesLabel.TabIndex = 8;
            this.ListofFilesLabel.Text = "List of Files";
            // 
            // listInstructionLabel
            // 
            this.listInstructionLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.listInstructionLabel.AutoSize = true;
            this.listInstructionLabel.Location = new System.Drawing.Point(16, 212);
            this.listInstructionLabel.Name = "listInstructionLabel";
            this.listInstructionLabel.Size = new System.Drawing.Size(243, 14);
            this.listInstructionLabel.TabIndex = 9;
            this.listInstructionLabel.Text = "Enter subject numbers separated by spaces";
            // 
            // SingleFileLabel
            // 
            this.SingleFileLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SingleFileLabel.AutoSize = true;
            this.SingleFileLabel.Font = new System.Drawing.Font("Calibri", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SingleFileLabel.Location = new System.Drawing.Point(273, 38);
            this.SingleFileLabel.Name = "SingleFileLabel";
            this.SingleFileLabel.Size = new System.Drawing.Size(122, 33);
            this.SingleFileLabel.TabIndex = 11;
            this.SingleFileLabel.Text = "Single File";
            // 
            // singleFileNum
            // 
            this.singleFileNum.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.singleFileNum.AutoSize = true;
            this.singleFileNum.Location = new System.Drawing.Point(372, 100);
            this.singleFileNum.Name = "singleFileNum";
            this.singleFileNum.Size = new System.Drawing.Size(120, 22);
            this.singleFileNum.TabIndex = 12;
            this.singleFileNum.ValueChanged += new System.EventHandler(this.singleFileNum_ValueChanged);
            // 
            // singleFileSelectLabel
            // 
            this.singleFileSelectLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.singleFileSelectLabel.AutoSize = true;
            this.singleFileSelectLabel.Location = new System.Drawing.Point(276, 102);
            this.singleFileSelectLabel.Name = "singleFileSelectLabel";
            this.singleFileSelectLabel.Size = new System.Drawing.Size(93, 14);
            this.singleFileSelectLabel.TabIndex = 13;
            this.singleFileSelectLabel.Text = "File to process: ";
            // 
            // listString
            // 
            this.listString.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.listString.Location = new System.Drawing.Point(19, 229);
            this.listString.Name = "listString";
            this.listString.Size = new System.Drawing.Size(212, 22);
            this.listString.TabIndex = 14;
            this.listString.Text = "ex. 0 1 3 4 5";
            this.listString.TextChanged += new System.EventHandler(this.listString_Changed);
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(279, 257);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(213, 23);
            this.progressBar.TabIndex = 15;
            // 
            // ProgressLabel
            // 
            this.ProgressLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ProgressLabel.AutoSize = true;
            this.ProgressLabel.Font = new System.Drawing.Font("Calibri", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProgressLabel.Location = new System.Drawing.Point(273, 170);
            this.ProgressLabel.Name = "ProgressLabel";
            this.ProgressLabel.Size = new System.Drawing.Size(109, 33);
            this.ProgressLabel.TabIndex = 16;
            this.ProgressLabel.Text = "Progress";
            // 
            // progressMessage1
            // 
            this.progressMessage1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.progressMessage1.AutoSize = true;
            this.progressMessage1.Location = new System.Drawing.Point(276, 212);
            this.progressMessage1.Name = "progressMessage1";
            this.progressMessage1.Size = new System.Drawing.Size(137, 14);
            this.progressMessage1.TabIndex = 17;
            this.progressMessage1.Text = "Current File: Not Started";
            // 
            // progressMessage2
            // 
            this.progressMessage2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.progressMessage2.AutoSize = true;
            this.progressMessage2.Location = new System.Drawing.Point(276, 232);
            this.progressMessage2.Name = "progressMessage2";
            this.progressMessage2.Size = new System.Drawing.Size(155, 14);
            this.progressMessage2.TabIndex = 18;
            this.progressMessage2.Text = "Current Section: Not Started";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menu});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(508, 24);
            this.menuStrip1.TabIndex = 19;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menu
            // 
            this.menu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menu_help,
            this.menu_about});
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(50, 20);
            this.menu.Text = "Menu";
            // 
            // menu_help
            // 
            this.menu_help.Name = "menu_help";
            this.menu_help.Size = new System.Drawing.Size(107, 22);
            this.menu_help.Text = "Help";
            this.menu_help.Click += new System.EventHandler(this.menu_help_Click);
            // 
            // menu_about
            // 
            this.menu_about.Name = "menu_about";
            this.menu_about.Size = new System.Drawing.Size(107, 22);
            this.menu_about.Text = "About";
            this.menu_about.Click += new System.EventHandler(this.menu_about_Click);
            // 
            // GUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 299);
            this.Controls.Add(this.progressMessage2);
            this.Controls.Add(this.progressMessage1);
            this.Controls.Add(this.ProgressLabel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.listString);
            this.Controls.Add(this.singleFileSelectLabel);
            this.Controls.Add(this.singleFileNum);
            this.Controls.Add(this.SingleFileLabel);
            this.Controls.Add(this.listInstructionLabel);
            this.Controls.Add(this.ListofFilesLabel);
            this.Controls.Add(this.maxFileNum);
            this.Controls.Add(this.maxFileNumLabel);
            this.Controls.Add(this.minFileNumLabel);
            this.Controls.Add(this.minFileNum);
            this.Controls.Add(this.RangeofFilesLabel);
            this.Controls.Add(this.singleStart);
            this.Controls.Add(this.listStart);
            this.Controls.Add(this.rangeStart);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "GUI";
            this.Text = "VisualTexcel";
            this.Load += new System.EventHandler(this.GUI_Load);
            ((System.ComponentModel.ISupportInitialize)(this.minFileNum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.maxFileNum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.singleFileNum)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button rangeStart;
        private System.Windows.Forms.Button listStart;
        private System.Windows.Forms.Button singleStart;

        private System.Windows.Forms.NumericUpDown minFileNum;
        private System.Windows.Forms.NumericUpDown maxFileNum;
        private System.Windows.Forms.NumericUpDown singleFileNum;

        private System.Windows.Forms.TextBox listString;

        private System.Windows.Forms.ProgressBar progressBar;

        private System.Windows.Forms.Label RangeofFilesLabel;
        private System.Windows.Forms.Label minFileNumLabel;
        private System.Windows.Forms.Label maxFileNumLabel;
        private System.Windows.Forms.Label ListofFilesLabel;
        private System.Windows.Forms.Label listInstructionLabel;
        private System.Windows.Forms.Label SingleFileLabel;
        private System.Windows.Forms.Label singleFileSelectLabel;
        private System.Windows.Forms.Label ProgressLabel;
        private System.Windows.Forms.Label progressMessage1;
        private System.Windows.Forms.Label progressMessage2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menu;
        private System.Windows.Forms.ToolStripMenuItem menu_help;
        private System.Windows.Forms.ToolStripMenuItem menu_about;
    }
}

