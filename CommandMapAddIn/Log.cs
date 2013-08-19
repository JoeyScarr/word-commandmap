using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace CommandMapAddIn {
	public static class Log {
		private static bool loggingEnabled = true;
		private static string filename = null;
		private static List<string> lines = new List<string>();
		private static int BUFFER_SIZE = 20;

		public static void DisableLogging() {
			loggingEnabled = false;
		}

		public static void StartLogging(string filepath) {
			if (filename != null) {
				Flush();
			}
			filename = filepath;
		}

		private static void FlushIfNecessary() {
			if (lines.Count >= BUFFER_SIZE) {
				Flush();
			}
		}

		public static void Flush() {
			if (loggingEnabled && filename != null) {
				Directory.CreateDirectory(Path.GetDirectoryName(filename));
				using (StreamWriter sw = new StreamWriter(filename, true)) {
					lock (lines) {
						foreach (string line in lines) {
							sw.WriteLine(line);
						}
						lines.Clear();
					}
				}
			} else {
				lines.Clear();
			}
		}

		public static void LogCommand(string msoName) {
			LogString(string.Format("COMMAND {0} {1}", msoName, DateTime.Now.Ticks));
		}

		public static void LogCommandMapOpen() {
			LogString(string.Format("CMOPEN {0}", DateTime.Now.Ticks));
		}

		public static void LogCommandMapClose() {
			LogString(string.Format("CMCLOSE {0}", DateTime.Now.Ticks));
		}

		public static void LogKeyDown(Keys key) {
			LogString(string.Format("CMOPEN {0} {1}", string.Concat("'", key, "'"), DateTime.Now.Ticks));
		}

		public static void LogString(string str) {
			lock (lines) {
				lines.Add(str);
			}
			FlushIfNecessary();
		}
	}
}
