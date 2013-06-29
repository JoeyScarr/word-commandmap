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

namespace WordCommandMap {
	public class WordInstance {
		Word.Application m_App;
		Word.Document m_Document;
		IntPtr m_WindowHandle;

		public WordInstance(string filename = null) {
			// Spawn the Word application.
			m_App = new Word.Application();
			m_App.Visible = true;
			m_App.Activate();
			
			// Either open the provided document, or start a new one.
			if (filename != null) {
				m_Document = m_App.Documents.Open(filename);
			} else {
				m_Document = m_App.Documents.Add();
			}
			m_WindowHandle = WindowsApi.GetForegroundWindow();
			Console.WriteLine("{0}: {1}",WindowsApi.GetWindowTitle(m_WindowHandle),WindowsApi.GetAppPath(m_WindowHandle));

			MinimizeRibbon();

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

		public void ColorPick() {
			m_Document.CommandBars.ExecuteMso("FontColorPicker");
		}
	}
}
