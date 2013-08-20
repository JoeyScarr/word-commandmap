using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CMStudy1 {
	public partial class StatusForm : Form {
		private int m_Participant = 0;
		private string m_App = "Pinta";
		private string m_Interface = "CM";
		private Process m_Process;

		private bool m_Running = false;

		public StatusForm() {
			InitializeComponent();
		}

		public StatusForm(int participant, string app, bool CM, Process process)
			: this() {
			m_Participant = participant;
			m_App = app;
			m_Process = process;
			m_Interface = CM ? "CM" : "Normal";
			Log.StartLogging(string.Format("logs\\P{0}_{1}_{2}.txt", m_Participant, m_App, m_Interface));
			Text = string.Format("P:{0} A:{1} I:{2}", m_Participant, m_App, m_Interface);
			UpdateStatus();

			// Set up a thread that will close this form when the app closes
			Thread t = new Thread(delegate() {
				m_Process.WaitForExit();
				this.Invoke(new Action(delegate() {
					Close();
				}));
			});
			t.Start();
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
				lStatus.Text = "Practice mode (not recording)";
				bStartStop.Text = "Click to Begin!";
			}
		}

		private void StatusForm_Load(object sender, EventArgs e) {
			Top = Screen.PrimaryScreen.WorkingArea.Height - Height - 30;
			Left = Screen.PrimaryScreen.WorkingArea.Width - Width - 30;
		}
	}
}
