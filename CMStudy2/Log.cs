using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace CMStudy2 {
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
			if (loggingEnabled) {
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

		public static void LogAppOpened() {
			LogString(string.Format("APPOPENED {0}", DateTime.Now.Ticks));
		}

		public static void LogTaskStart() {
			LogString(string.Format("TASKSTART {0}", DateTime.Now.Ticks));
		}

		public static void LogTaskEnd() {
			LogString(string.Format("TASKEND {0}", DateTime.Now.Ticks));
		}

		public static void LogAppClosed() {
			LogString(string.Format("APPCLOSED {0}", DateTime.Now.Ticks));
		}

		public static void LogString(string str) {
			lock (lines) {
				lines.Add(str);
			}
			FlushIfNecessary();
		}
	}
}
