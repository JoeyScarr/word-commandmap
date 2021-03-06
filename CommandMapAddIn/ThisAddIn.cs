﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using Gma.UserActivityMonitor;
using System.Threading;

namespace CommandMapAddIn {
	public partial class ThisAddIn {

		WordInstance m_Word;
		CommandMapForm m_CommandMap;
		ActivationButton m_ActivationButton;
		BlankRibbon m_BlankRibbon;
		NormalRibbon m_NormalRibbon;

		bool m_CtrlPressed = false;
		bool m_ShuttingDown = false;
		bool m_CommandMapEnabled = false;
		Keys m_CurrentKeyDown = Keys.None;

		private void ThisAddIn_Startup(object sender, System.EventArgs e) {
			m_CommandMapEnabled = GlobalSettings.GetCommandMapEnabled();
			// Inform the ribbon we're using about the current application
			if (m_BlankRibbon != null) {
				m_BlankRibbon.Application = Application;
			} else if (m_NormalRibbon != null) {
				m_NormalRibbon.Application = Application;
			}

			// Create a WordInstance
			m_Word = new WordInstance(Application);

#if LOGGING
			string logPath = GlobalSettings.GetLogPath();
			if (!string.IsNullOrEmpty(logPath)) {
				Log.StartLogging(logPath);
			}
#endif

			if (m_CommandMapEnabled) {
				// Spawn the on-screen activation button, and attach it to the Word window.
				m_ActivationButton = new ActivationButton(m_Word);
				m_ActivationButton.Click += m_ActivationButton_Click;
				m_ActivationButton.Show();

				// Spawn the CommandMap form, and attach it to the Word window.
				m_CommandMap = new CommandMapForm(m_Word);
				m_CommandMap.Show();
				m_CommandMap.Hide();
			}

			// Hook mouse events for logging
			HookManager.MouseDown += HookManager_MouseDown;
			HookManager.MouseUp += HookManager_MouseUp;

			// Hook keyboard events
			HookManager.KeyDown += HookManager_KeyDown;
			HookManager.KeyUp += HookManager_KeyUp;
		}

		void HookManager_MouseUp(object sender, MouseEventArgs e) {
			Thread t = new Thread(delegate() {
				Log.LogMouseUp(e);
			});
			t.Start();
		}

		void HookManager_MouseDown(object sender, MouseEventArgs e) {
			Thread t = new Thread(delegate() {
				if (!m_ShuttingDown) {
					Log.LogMouseDown(e);
				}
			});
			t.Start();
		}

		void m_ActivationButton_Click(object sender, EventArgs e) {
			m_CommandMap.Show();
		}

		void HookManager_KeyDown(object sender, KeyEventArgs e) {
			System.Action a = new System.Action(delegate() {
				if (m_CurrentKeyDown != e.KeyData) {
					Log.LogKeyDown(e.KeyData);
					m_CurrentKeyDown = e.KeyData;
				}
				if (m_CommandMapEnabled) {
					var key = e.KeyCode;
					if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey || key == Keys.Control) {
						if (!m_CommandMap.Visible && !m_CtrlPressed) {
							IntPtr foregroundWindow = WindowsApi.GetForegroundWindow();
							if (foregroundWindow == m_Word.WindowHandle
								|| foregroundWindow == m_CommandMap.Handle
								|| foregroundWindow == m_ActivationButton.Handle) {
								m_CtrlPressed = true;
								m_CommandMap.Show();
								Application.Activate();
							}
						}
					} else {
						m_CommandMap.Hide();
					}
				}
			});
			if (m_CommandMap != null) {
				m_CommandMap.BeginInvoke(a);
			} else {
				Thread t = new Thread(new ThreadStart(a));
				t.Start();
			}
		}

		void HookManager_KeyUp(object sender, KeyEventArgs e) {
			System.Action a = new System.Action(delegate() {
				m_CurrentKeyDown = Keys.None;
				if (m_CommandMapEnabled) {
					var key = e.KeyCode;
					if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey || key == Keys.Control) {
						m_CtrlPressed = false;
						m_CommandMap.Hide();
					}
				}
			});
			if (m_CommandMap != null) {
				m_CommandMap.BeginInvoke(a);
			} else {
				Thread t = new Thread(new ThreadStart(a));
				t.Start();
			}
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
			m_ShuttingDown = true;
			Log.Flush();
		}

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
			if (GlobalSettings.GetCommandMapEnabled()) {
				return (m_BlankRibbon = new BlankRibbon());
			} else {
				return (m_NormalRibbon = new NormalRibbon());
			}
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup() {
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
