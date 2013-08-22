using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CommandMapAddIn {
	public class WordInstance {
		Word.Application m_App;
		IntPtr m_WindowHandle;
		private HashSet<IntPtr> m_KnownChildren = new HashSet<IntPtr>();

		public WordInstance(Word.Application app) {
			m_App = app;
			m_App.Activate();
			
			// Get the window handle.
			m_WindowHandle = WindowsApi.GetForegroundWindow();
		}

		private void MinimizeRibbon() {
			if (m_App.CommandBars["Ribbon"].Height > 80) {
				var test = m_App.CommandBars["Ribbon"].Controls;
				// Not minimized, so toggle it
				m_App.ActiveWindow.ToggleRibbon();
			}
		}

		public Word.Application Application {
			get { return m_App; }
		}

		public IntPtr WindowHandle {
			get { return m_WindowHandle; }
		}

		public Rectangle GetWindowPosition() {
			return WindowsApi.GetWindowPosition(m_WindowHandle);
		}

		public void RegisterChild(IntPtr handle) {
			m_KnownChildren.Add(handle);
		}

		public void Focus() {
			// Check if there's a child window (e.g. a dialog box) that should be focused instead
			List<IntPtr> childWindows = WindowsApi.GetChildWindows(m_WindowHandle);
			foreach (IntPtr handle in childWindows) {
				if (!m_KnownChildren.Contains(handle)) {
					WindowsApi.SetForegroundWindow(handle);
					return;
				}
			}
			// If there isn't, just focus the main window
			m_App.Activate();
		}

		public void SendCommand(string p) {
			if (m_App.CommandBars.GetEnabledMso(p)) {
				try {
					m_App.CommandBars.ExecuteMso(p);
				} catch (COMException exception) {
					Debug.WriteLine(exception);
				}
			}
		}
	}
}
