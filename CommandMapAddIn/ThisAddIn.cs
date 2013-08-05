using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using Gma.UserActivityMonitor;

namespace CommandMapAddIn {
	public partial class ThisAddIn {

		WordInstance m_Word;
		CommandMapForm m_CommandMap;
		ActivationButton m_ActivationButton;
		BlankRibbon m_BlankRibbon;
		NormalRibbon m_NormalRibbon;

		bool m_CtrlPressed = false;

		private void ThisAddIn_Startup(object sender, System.EventArgs e) {
			// Inform the ribbon we're using about the current application
			if (m_BlankRibbon != null) {
				m_BlankRibbon.Application = Application;
			} else if (m_NormalRibbon != null) {
				m_NormalRibbon.Application = Application;
			}

			// Create a WordInstance
			m_Word = new WordInstance(Application);

			if (GlobalSettings.GetCommandMapEnabled()) {
				// Spawn the on-screen activation button, and attach it to the Word window.
				m_ActivationButton = new ActivationButton(m_Word);
				m_ActivationButton.Click += m_ActivationButton_Click;
				m_ActivationButton.Show();

				// Spawn the CommandMap form, and attach it to the Word window.
				m_CommandMap = new CommandMapForm(m_Word);

				// Add a global hook.
				HookManager.KeyDown += HookManager_KeyDown;
				HookManager.KeyUp += HookManager_KeyUp;
			}
		}

		void m_ActivationButton_Click(object sender, EventArgs e) {
			m_CommandMap.Show();
		}

		void HookManager_KeyDown(object sender, KeyEventArgs e) {
			var key = e.KeyCode;
			if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey || key == Keys.Control) {
				if (!m_CommandMap.Visible && !m_CtrlPressed && WindowsApi.GetForegroundWindow() == m_Word.WindowHandle) {
					m_CtrlPressed = true;
					m_CommandMap.Show();
					Application.Activate();
				}
			} else {
				m_CommandMap.Hide();
			}
		}

		void HookManager_KeyUp(object sender, KeyEventArgs e) {
			var key = e.KeyCode;
			if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey || key == Keys.Control) {
				m_CtrlPressed = false;
				m_CommandMap.Hide();
			}
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
			// Remove hooks.
			HookManager.ForceUnsubscribeFromGlobalKeyboardEvents();
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
