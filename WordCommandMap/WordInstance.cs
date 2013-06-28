using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordCommandMap {
	public class WordInstance {
		Word.Application m_App;
		IntPtr m_WindowHandle;

		public WordInstance(string filename = null) {
			// Spawn the Word application.
			m_App = new Word.Application();
			m_App.Visible = true;
			m_App.Activate();
			
			// Either open the provided document, or start a new one.
			if (filename != null) {
				m_App.Documents.Open(filename);
			} else {
				m_App.Documents.Add();
			}
			Thread.Sleep(1000);
			m_WindowHandle = WindowsApi.GetForegroundWindow();
			Console.WriteLine("{0}: {1}",WindowsApi.GetWindowTitle(m_WindowHandle),WindowsApi.GetAppPath(m_WindowHandle));
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
	}
}
