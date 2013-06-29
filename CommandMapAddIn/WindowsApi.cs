using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CommandMapAddIn {
	public static class WindowsApi {
		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
		[DllImport("user32.dll")]
		public static extern IntPtr GetActiveWindow();


		[StructLayout(LayoutKind.Sequential)]
		struct RECT {
			public RECT(Rectangle rectangle) {
				Left = rectangle.Left;
				Top = rectangle.Top;
				Right = rectangle.Right;
				Bottom = rectangle.Bottom;
			}
			public RECT(Point location, Size size) {
				Left = location.X;
				Top = location.Y;
				Right = location.X + size.Width;
				Bottom = location.Y + size.Height;
			}
			public Int32 Left;
			public Int32 Top;
			public Int32 Right;
			public Int32 Bottom;
		}
		[DllImport("user32.dll")]
		static extern bool GetWindowRect(IntPtr hwnd, ref RECT rect);
		[DllImport("user32.dll")]
		static extern bool GetClientRect(IntPtr hwnd, ref RECT rect);
		[DllImport("user32.dll")]
		static extern bool ClientToScreen(IntPtr hwnd, ref Point lpPoint);

		public static Rectangle GetWindowPosition(IntPtr handle) {
			RECT r = new RECT();
			if (GetClientRect(handle, ref r)) {
				Point topleft = new Point(r.Top, r.Left);
				ClientToScreen(handle, ref topleft);
				return new Rectangle(topleft.X, topleft.Y, r.Right - r.Left, r.Bottom - r.Top);
			} else {
				throw new Exception(string.Format("Can't get window position: Window {0} doesn't exist!", handle));
			}
			/*WINDOWPLACEMENT p = new WINDOWPLACEMENT();
			if (GetWindowPlacement(handle, ref p)) {
				return p.rcNormalPosition;
			} else {
				throw new Exception(string.Format("Can't get window position: Window {0} doesn't exist!", handle));
			}*/
		}

		[DllImport("user32.dll", EntryPoint = "GetWindowText", ExactSpelling = false, CharSet = CharSet.Auto)]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

		private const int MAXTITLE = 255;

		/// <summary>
		/// Get the title bar text of a window.
		/// </summary>
		/// <param name="hWnd">The handle of the window.</param>
		/// <returns>The name of the window.</returns>
		public static string GetWindowTitle(IntPtr hWnd) {
			StringBuilder title = new StringBuilder(MAXTITLE);
			int titleLength = GetWindowText(hWnd, title, title.Capacity + 1);
			title.Length = titleLength;

			return title.ToString();
		}

		[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
		static extern uint GetLongPathName(
				string lpszShortPath,
				[Out] StringBuilder lpszLongPath,
				uint cchBuffer);

		/// <summary>
		/// Convert the given file path to long form. This is to account
		/// for the occasional use of "short form" in Windows. If the file path
		/// is not in short form, the same string will be returned.
		/// </summary>
		/// <param name="shortName">The short name to convert.</param>
		/// <returns></returns>
		public static string ToLongPathName(string shortName) {
			StringBuilder longNameBuffer = new StringBuilder(256);
			uint bufferSize = (uint)longNameBuffer.Capacity;

			GetLongPathName(shortName, longNameBuffer, bufferSize);

			return longNameBuffer.ToString().ToLower();
		}


		[DllImport("user32.dll")]
		static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		/// <summary>
		/// Get the PID of a specified window.
		/// </summary>
		/// <param name="windowHandle">The handle of the window.</param>
		/// <returns>The PID of the window.</returns>
		public static uint GetProcessID(IntPtr windowHandle) {
			uint id;
			GetWindowThreadProcessId(windowHandle, out id);
			return id;
		}

		/// <summary>
		/// Retrieves the path of the app for a given window handle.
		/// </summary>
		public static string GetAppPath(IntPtr handle) {
			try {
				Process p = Process.GetProcessById((int)GetProcessID(handle));
				return ToLongPathName(p.MainModule.FileName);
			} catch (Exception e) {
				Console.WriteLine(e);
				return null;
			}
		}
	}
}
