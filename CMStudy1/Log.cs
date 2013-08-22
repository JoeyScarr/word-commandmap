using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace CMStudy1 {
	public static class Log {
		private static bool loggingEnabled = true;
		private static string filename = null;
		private static List<string> lines = new List<string>();
		private static int BUFFER_SIZE = 20;
		private static DateTime m_LastTaskStart;

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
				lock (lines) {
					using (StreamWriter sw = new StreamWriter(filename, true)) {
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

		public static void LogTaskStart() {
			m_LastTaskStart = DateTime.Now;
			LogString(string.Format("TASKSTART {0}", m_LastTaskStart.Ticks));
		}

		public static void LogTaskEnd() {
			DateTime taskEnd = DateTime.Now;
			LogString(string.Format("TASKEND {0} {1}", taskEnd.Ticks, (taskEnd - m_LastTaskStart).TotalSeconds));
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
