namespace DemoOpennessIntroduction
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
            this.ExcelFileName = new System.Windows.Forms.TextBox();
            this.SelectExcelFile = new System.Windows.Forms.Button();
            this.CreateProject = new System.Windows.Forms.Button();
            this.ProjectName = new System.Windows.Forms.TextBox();
            this.CopyFromLibrary = new System.Windows.Forms.Button();
            this.CopyToLibrary = new System.Windows.Forms.Button();
            this.GenerateHMIPage = new System.Windows.Forms.Button();
            this.Verbose = new System.Windows.Forms.TextBox();
            this.SkipCreate = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // ExcelFileName
            // 
            this.ExcelFileName.Location = new System.Drawing.Point(191, 30);
            this.ExcelFileName.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ExcelFileName.Name = "ExcelFileName";
            this.ExcelFileName.Size = new System.Drawing.Size(305, 20);
            this.ExcelFileName.TabIndex = 0;
            // 
            // SelectExcelFile
            // 
            this.SelectExcelFile.Location = new System.Drawing.Point(38, 25);
            this.SelectExcelFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SelectExcelFile.Name = "SelectExcelFile";
            this.SelectExcelFile.Size = new System.Drawing.Size(134, 26);
            this.SelectExcelFile.TabIndex = 1;
            this.SelectExcelFile.Text = "Seleziona File Excel";
            this.SelectExcelFile.UseVisualStyleBackColor = true;
            this.SelectExcelFile.Click += new System.EventHandler(this.SelectExcelFile_Click);
            // 
            // CreateProject
            // 
            this.CreateProject.Location = new System.Drawing.Point(93, 66);
            this.CreateProject.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CreateProject.Name = "CreateProject";
            this.CreateProject.Size = new System.Drawing.Size(100, 26);
            this.CreateProject.TabIndex = 2;
            this.CreateProject.Text = "Genera Progetto";
            this.CreateProject.UseVisualStyleBackColor = true;
            this.CreateProject.Click += new System.EventHandler(this.CreateProject_Click);
            // 
            // ProjectName
            // 
            this.ProjectName.Location = new System.Drawing.Point(196, 70);
            this.ProjectName.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ProjectName.Name = "ProjectName";
            this.ProjectName.Size = new System.Drawing.Size(140, 20);
            this.ProjectName.TabIndex = 3;
            this.ProjectName.Text = "projectname";
            // 
            // CopyFromLibrary
            // 
            this.CopyFromLibrary.Location = new System.Drawing.Point(49, 115);
            this.CopyFromLibrary.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CopyFromLibrary.Name = "CopyFromLibrary";
            this.CopyFromLibrary.Size = new System.Drawing.Size(200, 26);
            this.CopyFromLibrary.TabIndex = 4;
            this.CopyFromLibrary.Text = "Crea Software Da Libreria Globale";
            this.CopyFromLibrary.UseVisualStyleBackColor = true;
            this.CopyFromLibrary.Click += new System.EventHandler(this.CopyFromLibrary_Click);
            // 
            // CopyToLibrary
            // 
            this.CopyToLibrary.Location = new System.Drawing.Point(267, 115);
            this.CopyToLibrary.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.CopyToLibrary.Name = "CopyToLibrary";
            this.CopyToLibrary.Size = new System.Drawing.Size(200, 26);
            this.CopyToLibrary.TabIndex = 5;
            this.CopyToLibrary.Text = "Salva Software in Libreria Progetto";
            this.CopyToLibrary.UseVisualStyleBackColor = true;
            this.CopyToLibrary.Click += new System.EventHandler(this.CopyToLibrary_Click);
            // 
            // GenerateHMIPage
            // 
            this.GenerateHMIPage.Location = new System.Drawing.Point(191, 150);
            this.GenerateHMIPage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.GenerateHMIPage.Name = "GenerateHMIPage";
            this.GenerateHMIPage.Size = new System.Drawing.Size(134, 26);
            this.GenerateHMIPage.TabIndex = 6;
            this.GenerateHMIPage.Text = "Genera pagina Unified";
            this.GenerateHMIPage.UseVisualStyleBackColor = true;
            this.GenerateHMIPage.Click += new System.EventHandler(this.GenerateHMIPage_Click);
            // 
            // Verbose
            // 
            this.Verbose.Location = new System.Drawing.Point(70, 180);
            this.Verbose.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Verbose.Multiline = true;
            this.Verbose.Name = "Verbose";
            this.Verbose.Size = new System.Drawing.Size(385, 264);
            this.Verbose.TabIndex = 7;
            // 
            // SkipCreate
            // 
            this.SkipCreate.AutoSize = true;
            this.SkipCreate.ImageAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.SkipCreate.Location = new System.Drawing.Point(360, 72);
            this.SkipCreate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SkipCreate.Name = "SkipCreate";
            this.SkipCreate.Size = new System.Drawing.Size(100, 17);
            this.SkipCreate.TabIndex = 8;
            this.SkipCreate.Text = "Reopen Project";
            this.SkipCreate.UseVisualStyleBackColor = true;
            this.SkipCreate.CheckedChanged += new System.EventHandler(this.SkipCreate_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(533, 474);
            this.Controls.Add(this.SkipCreate);
            this.Controls.Add(this.Verbose);
            this.Controls.Add(this.GenerateHMIPage);
            this.Controls.Add(this.CopyToLibrary);
            this.Controls.Add(this.CopyFromLibrary);
            this.Controls.Add(this.ProjectName);
            this.Controls.Add(this.CreateProject);
            this.Controls.Add(this.SelectExcelFile);
            this.Controls.Add(this.ExcelFileName);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox ExcelFileName;
        private System.Windows.Forms.Button SelectExcelFile;
        private System.Windows.Forms.Button CreateProject;
        private System.Windows.Forms.TextBox ProjectName;
        private System.Windows.Forms.Button CopyFromLibrary;
        private System.Windows.Forms.Button CopyToLibrary;
        private System.Windows.Forms.Button GenerateHMIPage;
        private System.Windows.Forms.TextBox Verbose;
        private System.Windows.Forms.CheckBox SkipCreate;
    }
}

