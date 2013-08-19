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
    private static string mouseFilename = null;
    private static List<string> lines = new List<string>();
    private static List<string> mouseLines = new List<string>();
    private static int BUFFER_SIZE = 20;

    public static void DisableLogging() {
      loggingEnabled = false;
    }

    public static void StartLogging(string filepath, string mouseFilePath = null) {
      if (filename != null) {
        Flush();
      }
      filename = filepath;
      if (mouseFilename != null) {
        FlushMouse();
      }
      mouseFilename = mouseFilePath;
    }

    private static void FlushIfNecessary() {
      if (lines.Count >= BUFFER_SIZE) {
        Flush();
      }
			if (mouseFilename != null && mouseLines.Count >= BUFFER_SIZE) {
        FlushMouse();
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

    public static void FlushMouse() {
      if (loggingEnabled) {
        using (StreamWriter sw = new StreamWriter(mouseFilename, true)) {
          lock (mouseLines) {
            foreach (string line in mouseLines) {
              sw.WriteLine(line);
            }
            mouseLines.Clear();
          }
        }
      } else {
        mouseLines.Clear();
      }
    }

		public static void LogTaskStart() {
			LogString(string.Format("TASKSTART {0}", DateTime.Now.Ticks));
		}

		public static void LogTaskEnd() {
			LogString(string.Format("TASKEND {0}", DateTime.Now.Ticks));
		}

    public static void LogString(string str) {
      lock (lines) {
        lines.Add(str);
      }
      FlushIfNecessary();
    }

    public static void LogMousePosition(Point pos) {
      lock (mouseLines) {
        mouseLines.Add(
          string.Format("{0} {1} {2}", DateTime.Now.Ticks, pos.X, pos.Y));
      }
    }
  }
}
