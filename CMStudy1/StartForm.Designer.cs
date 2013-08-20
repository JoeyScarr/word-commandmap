namespace CMStudy1 {
	partial class StartForm {
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
			this.label1 = new System.Windows.Forms.Label();
			this.numParticipant = new System.Windows.Forms.NumericUpDown();
			this.bStartWordCM = new System.Windows.Forms.Button();
			this.bStartWordNormal = new System.Windows.Forms.Button();
			this.bStartPracticeNormal = new System.Windows.Forms.Button();
			this.bStartPracticeCM = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.numParticipant)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(98, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Participant number:";
			// 
			// numParticipant
			// 
			this.numParticipant.Location = new System.Drawing.Point(116, 7);
			this.numParticipant.Maximum = new decimal(new int[] {
            24,
            0,
            0,
            0});
			this.numParticipant.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numParticipant.Name = "numParticipant";
			this.numParticipant.ReadOnly = true;
			this.numParticipant.Size = new System.Drawing.Size(62, 20);
			this.numParticipant.TabIndex = 1;
			this.numParticipant.TabStop = false;
			this.numParticipant.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numParticipant.ValueChanged += new System.EventHandler(this.numParticipant_ValueChanged);
			// 
			// bStartWordCM
			// 
			this.bStartWordCM.Location = new System.Drawing.Point(15, 33);
			this.bStartWordCM.Name = "bStartWordCM";
			this.bStartWordCM.Size = new System.Drawing.Size(166, 37);
			this.bStartWordCM.TabIndex = 4;
			this.bStartWordCM.Text = "CommandMap";
			this.bStartWordCM.UseVisualStyleBackColor = true;
			this.bStartWordCM.Click += new System.EventHandler(this.bStartWordCM_Click);
			// 
			// bStartWordNormal
			// 
			this.bStartWordNormal.Location = new System.Drawing.Point(15, 76);
			this.bStartWordNormal.Name = "bStartWordNormal";
			this.bStartWordNormal.Size = new System.Drawing.Size(166, 37);
			this.bStartWordNormal.TabIndex = 5;
			this.bStartWordNormal.Text = "Normal";
			this.bStartWordNormal.UseVisualStyleBackColor = true;
			this.bStartWordNormal.Click += new System.EventHandler(this.bStartWordNormal_Click);
			// 
			// bStartPracticeNormal
			// 
			this.bStartPracticeNormal.Location = new System.Drawing.Point(15, 163);
			this.bStartPracticeNormal.Name = "bStartPracticeNormal";
			this.bStartPracticeNormal.Size = new System.Drawing.Size(166, 37);
			this.bStartPracticeNormal.TabIndex = 7;
			this.bStartPracticeNormal.Text = "Normal (PRACTICE)";
			this.bStartPracticeNormal.UseVisualStyleBackColor = true;
			this.bStartPracticeNormal.Click += new System.EventHandler(this.bStartPracticeNormal_Click);
			// 
			// bStartPracticeCM
			// 
			this.bStartPracticeCM.Location = new System.Drawing.Point(15, 120);
			this.bStartPracticeCM.Name = "bStartPracticeCM";
			this.bStartPracticeCM.Size = new System.Drawing.Size(166, 37);
			this.bStartPracticeCM.TabIndex = 6;
			this.bStartPracticeCM.Text = "CommandMap (PRACTICE)";
			this.bStartPracticeCM.UseVisualStyleBackColor = true;
			this.bStartPracticeCM.Click += new System.EventHandler(this.bStartPracticeCM_Click);
			// 
			// StartForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(190, 215);
			this.Controls.Add(this.bStartPracticeNormal);
			this.Controls.Add(this.bStartPracticeCM);
			this.Controls.Add(this.bStartWordNormal);
			this.Controls.Add(this.bStartWordCM);
			this.Controls.Add(this.numParticipant);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "StartForm";
			this.Text = "Experiment";
			((System.ComponentModel.ISupportInitialize)(this.numParticipant)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.NumericUpDown numParticipant;
		private System.Windows.Forms.Button bStartWordCM;
		private System.Windows.Forms.Button bStartWordNormal;
		private System.Windows.Forms.Button bStartPracticeNormal;
		private System.Windows.Forms.Button bStartPracticeCM;
	}
}