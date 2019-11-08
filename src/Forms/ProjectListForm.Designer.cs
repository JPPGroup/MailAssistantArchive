﻿namespace Jpp.AddIn.MailAssistant.Forms
{
    partial class ProjectListForm
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.txtSearchBox = new System.Windows.Forms.TextBox();
            this.panMain = new System.Windows.Forms.Panel();
            this.lstProjects = new System.Windows.Forms.ListView();
            this.lblSearch = new System.Windows.Forms.Label();
            this.panMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(691, 373);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(610, 373);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // txtSearchBox
            // 
            this.txtSearchBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSearchBox.Location = new System.Drawing.Point(64, 12);
            this.txtSearchBox.Name = "txtSearchBox";
            this.txtSearchBox.Size = new System.Drawing.Size(674, 20);
            this.txtSearchBox.TabIndex = 3;
            this.txtSearchBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TxtSearchBox_KeyUp);
            // 
            // panMain
            // 
            this.panMain.BackColor = System.Drawing.Color.White;
            this.panMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panMain.Controls.Add(this.lblSearch);
            this.panMain.Controls.Add(this.lstProjects);
            this.panMain.Controls.Add(this.txtSearchBox);
            this.panMain.Location = new System.Drawing.Point(12, 12);
            this.panMain.Name = "panMain";
            this.panMain.Size = new System.Drawing.Size(754, 355);
            this.panMain.TabIndex = 4;
            // 
            // lstProjects
            // 
            this.lstProjects.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstProjects.FullRowSelect = true;
            this.lstProjects.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstProjects.HideSelection = false;
            this.lstProjects.Location = new System.Drawing.Point(14, 38);
            this.lstProjects.MultiSelect = false;
            this.lstProjects.Name = "lstProjects";
            this.lstProjects.Size = new System.Drawing.Size(724, 300);
            this.lstProjects.TabIndex = 4;
            this.lstProjects.UseCompatibleStateImageBehavior = false;
            // 
            // lblSearch
            // 
            this.lblSearch.AutoSize = true;
            this.lblSearch.Location = new System.Drawing.Point(11, 14);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new System.Drawing.Size(47, 13);
            this.lblSearch.TabIndex = 5;
            this.lblSearch.Text = "Search :";
            // 
            // ProjectListForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(778, 408);
            this.ControlBox = false;
            this.Controls.Add(this.panMain);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProjectListForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select a project...";
            this.Load += new System.EventHandler(this.ProjectListForm_Load);
            this.panMain.ResumeLayout(false);
            this.panMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TextBox txtSearchBox;
        private System.Windows.Forms.Panel panMain;
        private System.Windows.Forms.ListView lstProjects;
        private System.Windows.Forms.Label lblSearch;
    }
}