using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace CommandMapAddIn {
	public static class GlobalSettings {
		public const int TITLEBAR_HEIGHT = 55;
		public const int BASE_RIBBON_HEIGHT = 93;
		public const int CM_RIBBON_HEIGHT = 118;
		public const int NUM_CM_RIBBONS = 6;
		public const int STATUSBAR_HEIGHT = 22;

		public static bool GetCommandMapEnabled() {
			RegistryKey key = Registry.CurrentUser.CreateSubKey("WordCommandMap");
			var val = (Int32)key.GetValue("CMEnabled", 1);
			return val != 0;
		}

		public static void SetCommandMapEnabled(bool value) {
			RegistryKey key = Registry.CurrentUser.CreateSubKey("WordCommandMap");
			key.SetValue("CMEnabled", value ? 1 : 0, RegistryValueKind.DWord);
			key.Close();
		}

		public static string GetLogPath() {
			RegistryKey key = Registry.CurrentUser.CreateSubKey("WordCommandMap");
			string val = (string)key.GetValue("LogPath", null);
			return val;
		}
	}
}
