namespace WordCommandMap {
    partial class MainForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.button1 = new System.Windows.Forms.Button();
            this.ribbon1 = new System.Windows.Forms.Ribbon();
            this.ribbonTab1 = new System.Windows.Forms.RibbonTab();
            this.ribbonPanel1 = new System.Windows.Forms.RibbonPanel();
            this.ribbonPanel2 = new System.Windows.Forms.RibbonPanel();
            this.ribbon2 = new System.Windows.Forms.Ribbon();
            this.ribbonTab2 = new System.Windows.Forms.RibbonTab();
            this.ribbonPanel3 = new System.Windows.Forms.RibbonPanel();
            this.ribbonComboBox1 = new System.Windows.Forms.RibbonComboBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(257, 364);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(327, 171);
            this.button1.TabIndex = 0;
            this.button1.Text = "Do something in Word";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // ribbon1
            // 
            this.ribbon1.CaptionBarVisible = false;
            this.ribbon1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.ribbon1.Location = new System.Drawing.Point(0, 0);
            this.ribbon1.Minimized = false;
            this.ribbon1.Name = "ribbon1";
            // 
            // 
            // 
            this.ribbon1.OrbDropDown.BorderRoundness = 8;
            this.ribbon1.OrbDropDown.Location = new System.Drawing.Point(0, 0);
            this.ribbon1.OrbDropDown.Name = "";
            this.ribbon1.OrbDropDown.Size = new System.Drawing.Size(527, 447);
            this.ribbon1.OrbDropDown.TabIndex = 0;
            this.ribbon1.OrbImage = null;
            this.ribbon1.Size = new System.Drawing.Size(713, 127);
            this.ribbon1.TabIndex = 1;
            this.ribbon1.Tabs.Add(this.ribbonTab1);
            this.ribbon1.TabsMargin = new System.Windows.Forms.Padding(12, 2, 20, 0);
            this.ribbon1.Text = "ribbon1";
            // 
            // ribbonTab1
            // 
            this.ribbonTab1.Panels.Add(this.ribbonPanel1);
            this.ribbonTab1.Panels.Add(this.ribbonPanel2);
            this.ribbonTab1.Text = "ribbonTab1";
            // 
            // ribbonPanel1
            // 
            this.ribbonPanel1.Text = "Clipboard";
            // 
            // ribbonPanel2
            // 
            this.ribbonPanel2.Text = "Font";
            // 
            // ribbon2
            // 
            this.ribbon2.CaptionBarVisible = false;
            this.ribbon2.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.ribbon2.Location = new System.Drawing.Point(0, 127);
            this.ribbon2.Minimized = false;
            this.ribbon2.Name = "ribbon2";
            // 
            // 
            // 
            this.ribbon2.OrbDropDown.BorderRoundness = 8;
            this.ribbon2.OrbDropDown.Location = new System.Drawing.Point(0, 0);
            this.ribbon2.OrbDropDown.Name = "";
            this.ribbon2.OrbDropDown.Size = new System.Drawing.Size(527, 447);
            this.ribbon2.OrbDropDown.TabIndex = 0;
            this.ribbon2.OrbImage = null;
            this.ribbon2.OrbVisible = false;
            // 
            // 
            // 
            this.ribbon2.QuickAcessToolbar.Visible = false;
            this.ribbon2.Size = new System.Drawing.Size(713, 200);
            this.ribbon2.TabIndex = 2;
            this.ribbon2.Tabs.Add(this.ribbonTab2);
            this.ribbon2.TabsMargin = new System.Windows.Forms.Padding(12, 2, 20, 0);
            this.ribbon2.Text = "ribbon2";
            // 
            // ribbonTab2
            // 
            this.ribbonTab2.Panels.Add(this.ribbonPanel3);
            this.ribbonTab2.Text = "ribbonTab2";
            // 
            // ribbonPanel3
            // 
            this.ribbonPanel3.Text = "ribbonPanel3";
            // 
            // ribbonComboBox1
            // 
            this.ribbonComboBox1.AllowTextEdit = false;
            this.ribbonComboBox1.DrawIconsBar = false;
            this.ribbonComboBox1.DropDownResizable = true;
            this.ribbonComboBox1.Text = "Date:";
            this.ribbonComboBox1.TextBoxText = "";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(713, 547);
            this.Controls.Add(this.ribbon2);
            this.Controls.Add(this.ribbon1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "MainForm";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Ribbon ribbon1;
        private System.Windows.Forms.RibbonTab ribbonTab1;
        private System.Windows.Forms.RibbonPanel ribbonPanel1;
        private System.Windows.Forms.RibbonPanel ribbonPanel2;
        private System.Windows.Forms.Ribbon ribbon2;
        private System.Windows.Forms.RibbonTab ribbonTab2;
        private System.Windows.Forms.RibbonPanel ribbonPanel3;
        private System.Windows.Forms.RibbonComboBox ribbonComboBox1;
    }
}

