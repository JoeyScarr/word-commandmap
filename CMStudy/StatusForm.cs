﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CMStudy {
	public partial class StatusForm : Form {
		private int m_Participant = 0;
		private int m_Block = 0;
		private string m_App = "Pinta";
		private string m_Interface = "CM";

		private bool m_Running = false;

		public StatusForm() {
			InitializeComponent();
		}

		public StatusForm(int participant, int block, string app, bool CM)
			: this() {
			m_Participant = participant;
			m_Block = block;
			m_App = app;
			m_Interface = CM ? "CM" : "Normal";
			Log.StartLogging(string.Format("logs\\P{0}_D{1}_{2}_{3}.txt", m_Participant, m_Block, m_App, m_Interface));
			Text = string.Format("P:{0} D:{1} A:{2} I:{3}", m_Participant, m_Block, m_App, m_Interface);
		}

		private void bStartStop_Click(object sender, EventArgs e) {
			if (!m_Running) {
				Log.LogTaskStart();
			} else {
				Log.LogTaskEnd();
				Log.Flush();
			}
			m_Running = !m_Running;
			UpdateStatus();
		}

		private void UpdateStatus() {
			if (m_Running) {
				lStatus.Text = "Recording started at " + DateTime.Now.ToShortTimeString();
				bStartStop.Text = "Click when Finished";
			} else {
				lStatus.Text = "Not recording";
				bStartStop.Text = "Click to Begin!";
			}
		}
	}
}
