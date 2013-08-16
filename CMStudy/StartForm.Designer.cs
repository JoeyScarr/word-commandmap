namespace CMStudy {
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
			this.numDay = new System.Windows.Forms.NumericUpDown();
			this.label2 = new System.Windows.Forms.Label();
			this.bStartWordCM = new System.Windows.Forms.Button();
			this.bStartWordNormal = new System.Windows.Forms.Button();
			this.bStartPintaCM = new System.Windows.Forms.Button();
			this.bStartPintaNormal = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.numParticipant)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.numDay)).BeginInit();
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
			// numDay
			// 
			this.numDay.Location = new System.Drawing.Point(116, 33);
			this.numDay.Maximum = new decimal(new int[] {
            5,
            0,
            0,
            0});
			this.numDay.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.numDay.Name = "numDay";
			this.numDay.Size = new System.Drawing.Size(62, 20);
			this.numDay.TabIndex = 3;
			this.numDay.TabStop = false;
			this.numDay.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 35);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 13);
			this.label2.TabIndex = 2;
			this.label2.Text = "Day:";
			// 
			// bStartWordCM
			// 
			this.bStartWordCM.Location = new System.Drawing.Point(12, 59);
			this.bStartWordCM.Name = "bStartWordCM";
			this.bStartWordCM.Size = new System.Drawing.Size(166, 37);
			this.bStartWordCM.TabIndex = 4;
			this.bStartWordCM.Text = "Start Word (CM)";
			this.bStartWordCM.UseVisualStyleBackColor = true;
			this.bStartWordCM.Click += new System.EventHandler(this.bStartWordCM_Click);
			// 
			// bStartWordNormal
			// 
			this.bStartWordNormal.Location = new System.Drawing.Point(12, 102);
			this.bStartWordNormal.Name = "bStartWordNormal";
			this.bStartWordNormal.Size = new System.Drawing.Size(166, 37);
			this.bStartWordNormal.TabIndex = 5;
			this.bStartWordNormal.Text = "Start Word (Normal)";
			this.bStartWordNormal.UseVisualStyleBackColor = true;
			this.bStartWordNormal.Click += new System.EventHandler(this.bStartWordNormal_Click);
			// 
			// bStartPintaCM
			// 
			this.bStartPintaCM.Location = new System.Drawing.Point(12, 145);
			this.bStartPintaCM.Name = "bStartPintaCM";
			this.bStartPintaCM.Size = new System.Drawing.Size(166, 37);
			this.bStartPintaCM.TabIndex = 5;
			this.bStartPintaCM.Text = "Start Pinta (CM)";
			this.bStartPintaCM.UseVisualStyleBackColor = true;
			this.bStartPintaCM.Click += new System.EventHandler(this.bStartPintaCM_Click);
			// 
			// bStartPintaNormal
			// 
			this.bStartPintaNormal.Location = new System.Drawing.Point(12, 188);
			this.bStartPintaNormal.Name = "bStartPintaNormal";
			this.bStartPintaNormal.Size = new System.Drawing.Size(166, 37);
			this.bStartPintaNormal.TabIndex = 5;
			this.bStartPintaNormal.Text = "Start Pinta (Normal)";
			this.bStartPintaNormal.UseVisualStyleBackColor = true;
			this.bStartPintaNormal.Click += new System.EventHandler(this.bStartPintaNormal_Click);
			// 
			// StartForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(190, 238);
			this.Controls.Add(this.bStartPintaNormal);
			this.Controls.Add(this.bStartPintaCM);
			this.Controls.Add(this.bStartWordNormal);
			this.Controls.Add(this.bStartWordCM);
			this.Controls.Add(this.numDay);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.numParticipant);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "StartForm";
			this.Text = "Experiment";
			((System.ComponentModel.ISupportInitialize)(this.numParticipant)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.numDay)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.NumericUpDown numParticipant;
		private System.Windows.Forms.NumericUpDown numDay;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button bStartWordCM;
		private System.Windows.Forms.Button bStartWordNormal;
		private System.Windows.Forms.Button bStartPintaCM;
		private System.Windows.Forms.Button bStartPintaNormal;
	}
}