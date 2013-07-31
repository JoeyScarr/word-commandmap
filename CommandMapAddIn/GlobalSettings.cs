using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace CommandMapAddIn {
	public static class GlobalSettings {
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
	}
}
