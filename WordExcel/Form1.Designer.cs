﻿namespace WordExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SelectWordFileButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SelectWordFileButton
            // 
            this.SelectWordFileButton.Location = new System.Drawing.Point(239, 205);
            this.SelectWordFileButton.Name = "SelectWordFileButton";
            this.SelectWordFileButton.Size = new System.Drawing.Size(75, 23);
            this.SelectWordFileButton.TabIndex = 0;
            this.SelectWordFileButton.Text = "button1";
            this.SelectWordFileButton.UseVisualStyleBackColor = true;
            this.SelectWordFileButton.Click += new System.EventHandler(this.SelectWordFileButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.SelectWordFileButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private Button SelectWordFileButton;
    }
}