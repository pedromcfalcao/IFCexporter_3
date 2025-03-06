namespace IFCexporter_3
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
            this.textResult = new System.Windows.Forms.TextBox();
            this.Update = new System.Windows.Forms.Button();
            this.Refresh = new System.Windows.Forms.Button();
            this.IFC_Export = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textResult
            // 
            this.textResult.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.textResult.Location = new System.Drawing.Point(32, 51);
            this.textResult.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textResult.Multiline = true;
            this.textResult.Name = "textResult";
            this.textResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textResult.Size = new System.Drawing.Size(1114, 538);
            this.textResult.TabIndex = 12;
            // 
            // Update
            // 
            this.Update.Location = new System.Drawing.Point(132, 602);
            this.Update.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Update.Name = "Update";
            this.Update.Size = new System.Drawing.Size(98, 78);
            this.Update.TabIndex = 19;
            this.Update.Text = "Update Model";
            this.Update.UseVisualStyleBackColor = true;
            this.Update.Click += new System.EventHandler(this.Update_Click);
            // 
            // Refresh
            // 
            this.Refresh.Location = new System.Drawing.Point(28, 602);
            this.Refresh.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Refresh.Name = "Refresh";
            this.Refresh.Size = new System.Drawing.Size(88, 80);
            this.Refresh.TabIndex = 20;
            this.Refresh.Text = "Refresh";
            this.Refresh.UseVisualStyleBackColor = true;
            this.Refresh.Click += new System.EventHandler(this.Refresh_Click);
            // 
            // IFC_Export
            // 
            this.IFC_Export.Location = new System.Drawing.Point(244, 602);
            this.IFC_Export.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.IFC_Export.Name = "IFC_Export";
            this.IFC_Export.Size = new System.Drawing.Size(110, 78);
            this.IFC_Export.TabIndex = 21;
            this.IFC_Export.Text = "IFC Export";
            this.IFC_Export.UseVisualStyleBackColor = true;
            this.IFC_Export.Click += new System.EventHandler(this.IFC_Export_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1160, 714);
            this.Controls.Add(this.IFC_Export);
            this.Controls.Add(this.Refresh);
            this.Controls.Add(this.Update);
            this.Controls.Add(this.textResult);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Name = "Form1";
            this.Text = "IFC exporter";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textResult;
        private System.Windows.Forms.Button Update;
        private System.Windows.Forms.Button Refresh;
        private System.Windows.Forms.Button IFC_Export;
    }
}