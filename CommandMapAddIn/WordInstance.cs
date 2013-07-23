using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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

		public WordInstance(Word.Application app) {
			m_App = app;
			m_App.Activate();
			
			// Get the window handle.
			m_WindowHandle = WindowsApi.GetForegroundWindow();
			Console.WriteLine("{0}: {1}",WindowsApi.GetWindowTitle(m_WindowHandle),WindowsApi.GetAppPath(m_WindowHandle));
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

		public void SendCommand(string p) {
			if (m_App.CommandBars.GetEnabledMso(p)) {
				m_App.CommandBars.ExecuteMso(p);
			}
		}
	}
}
