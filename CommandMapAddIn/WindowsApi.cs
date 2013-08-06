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

		enum WindowLongFlags : int {
			GWL_EXSTYLE = -20,
			GWLP_HINSTANCE = -6,
			GWLP_HWNDPARENT = -8,
			GWL_ID = -12,
			GWL_STYLE = -16,
			GWL_USERDATA = -21,
			GWL_WNDPROC = -4,
			DWLP_USER = 0x8,
			DWLP_MSGRESULT = 0x0,
			DWLP_DLGPROC = 0x4
		}

		/// <summary>
		/// Window Styles.
		/// The following styles can be specified wherever a window style is required. After the control has been created, these styles cannot be modified, except as noted.
		/// </summary>
		[Flags()]
		private enum WindowStyles : uint {
			/// <summary>The window has a thin-line border.</summary>
			WS_BORDER = 0x800000,

			/// <summary>The window has a title bar (includes the WS_BORDER style).</summary>
			WS_CAPTION = 0xc00000,

			/// <summary>The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.</summary>
			WS_CHILD = 0x40000000,

			/// <summary>Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.</summary>
			WS_CLIPCHILDREN = 0x2000000,

			/// <summary>
			/// Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated.
			/// If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
			/// </summary>
			WS_CLIPSIBLINGS = 0x4000000,

			/// <summary>The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.</summary>
			WS_DISABLED = 0x8000000,

			/// <summary>The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.</summary>
			WS_DLGFRAME = 0x400000,

			/// <summary>
			/// The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style.
			/// The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
			/// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
			/// </summary>
			WS_GROUP = 0x20000,

			/// <summary>The window has a horizontal scroll bar.</summary>
			WS_HSCROLL = 0x100000,

			/// <summary>The window is initially maximized.</summary> 
			WS_MAXIMIZE = 0x1000000,

			/// <summary>The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary> 
			WS_MAXIMIZEBOX = 0x10000,

			/// <summary>The window is initially minimized.</summary>
			WS_MINIMIZE = 0x20000000,

			/// <summary>The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary>
			WS_MINIMIZEBOX = 0x20000,

			/// <summary>The window is an overlapped window. An overlapped window has a title bar and a border.</summary>
			WS_OVERLAPPED = 0x0,

			/// <summary>The window is an overlapped window.</summary>
			WS_OVERLAPPEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_SIZEFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,

			/// <summary>The window is a pop-up window. This style cannot be used with the WS_CHILD style.</summary>
			WS_POPUP = 0x80000000u,

			/// <summary>The window is a pop-up window. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.</summary>
			WS_POPUPWINDOW = WS_POPUP | WS_BORDER | WS_SYSMENU,

			/// <summary>The window has a sizing border.</summary>
			WS_SIZEFRAME = 0x40000,

			/// <summary>The window has a window menu on its title bar. The WS_CAPTION style must also be specified.</summary>
			WS_SYSMENU = 0x80000,

			/// <summary>
			/// The window is a control that can receive the keyboard focus when the user presses the TAB key.
			/// Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.  
			/// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
			/// For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.
			/// </summary>
			WS_TABSTOP = 0x10000,

			/// <summary>The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.</summary>
			WS_VISIBLE = 0x10000000,

			/// <summary>The window has a vertical scroll bar.</summary>
			WS_VSCROLL = 0x200000
		}
		
		[DllImport("user32.dll")]
		static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

		[DllImport("user32.dll", SetLastError = true)]
		static extern int GetWindowLong(IntPtr hWnd, int nIndex);

		[DllImport("user32.dll")]
		static extern IntPtr SetParent(IntPtr child, IntPtr parent);

		public static void SetWindowParent(IntPtr child, IntPtr parent) {
			/*SetParent(child, parent);
			int style = GetWindowLong(child, (int)WindowLongFlags.GWL_STYLE);
			SetWindowLong(child, (int)WindowLongFlags.GWL_STYLE, style | (int)(WindowStyles.WS_CHILD | WindowStyles.WS_VISIBLE));*/

			SetWindowLong(child, (int)WindowLongFlags.GWLP_HWNDPARENT, parent.ToInt32());
		}

		[DllImport("user32.dll")]
		public static extern bool ShowWindow(IntPtr handle, int flags);
	}
}
