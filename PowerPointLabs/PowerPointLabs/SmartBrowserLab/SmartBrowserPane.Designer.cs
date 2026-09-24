namespace PowerPointLabs.SmartBrowserLab
{
    partial class SmartBrowserPane
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.smartBrowserPaneWPF1 = new PowerPointLabs.SmartBrowserLab.Views.SmartBrowserPaneWPF();
            this.SuspendLayout();

            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(300, 833);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Child = this.smartBrowserPaneWPF1;

            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHost1);
            this.Name = "SmartBrowserPane";
            this.Size = new System.Drawing.Size(300, 833);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private PowerPointLabs.SmartBrowserLab.Views.SmartBrowserPaneWPF smartBrowserPaneWPF1;
    }
}
