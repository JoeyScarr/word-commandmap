namespace CMStudy2 {
	partial class StatusForm {
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
			this.lStatus = new System.Windows.Forms.Label();
			this.bStartStop = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lStatus
			// 
			this.lStatus.AutoSize = true;
			this.lStatus.Location = new System.Drawing.Point(12, 9);
			this.lStatus.Name = "lStatus";
			this.lStatus.Size = new System.Drawing.Size(62, 13);
			this.lStatus.TabIndex = 0;
			this.lStatus.Text = "Not running";
			// 
			// bStartStop
			// 
			this.bStartStop.Location = new System.Drawing.Point(12, 25);
			this.bStartStop.Name = "bStartStop";
			this.bStartStop.Size = new System.Drawing.Size(158, 35);
			this.bStartStop.TabIndex = 1;
			this.bStartStop.Text = "Start";
			this.bStartStop.UseVisualStyleBackColor = true;
			this.bStartStop.Click += new System.EventHandler(this.bStartStop_Click);
			// 
			// StatusForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(182, 72);
			this.Controls.Add(this.bStartStop);
			this.Controls.Add(this.lStatus);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "StatusForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Stopped";
			this.TopMost = true;
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lStatus;
		private System.Windows.Forms.Button bStartStop;
	}
}

