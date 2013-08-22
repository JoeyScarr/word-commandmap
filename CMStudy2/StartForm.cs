using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CMStudy2 {
	public partial class StartForm : Form {
		public StartForm() {
			InitializeComponent();

			UpdateOrdering();
		}

		private void numParticipant_ValueChanged(object sender, EventArgs e) {
			UpdateOrdering();
		}

		private void numParticipant_KeyDown(object sender, KeyEventArgs e) {
			UpdateOrdering();
		}

		private void UpdateOrdering() {
			int participant = (int)numParticipant.Value - 1;

			int idxCMP = participant % 2 * 2 + (participant / 2) % 2;
			int idxCMW = participant % 2 * 2 + (participant / 2 + 1) % 2;
			int idxNP = (participant + 1) % 2 * 2 + (participant / 2) % 2;
			int idxNW = (participant + 1) % 2 * 2 + (participant / 2 + 1) % 2;

			bStartPintaCM.Top = 59 + idxCMP * 43;
			bStartWordCM.Top = 59 + idxCMW * 43;
			bStartPintaNormal.Top = 59 + idxNP * 43;
			bStartWordNormal.Top = 59 + idxNW * 43;

			bStartPintaCM.Enabled = true;
			bStartWordCM.Enabled = true;
			bStartPintaNormal.Enabled = true;
			bStartWordNormal.Enabled = true;
		}

		private void bStartWordCM_Click(object sender, EventArgs e) {
			StartWord2007(CM: true);
			bStartWordCM.Enabled = false;
		}

		private void bStartWordNormal_Click(object sender, EventArgs e) {
			StartWord2007(CM: false);
			bStartWordNormal.Enabled = false;
		}

		private void bStartPintaCM_Click(object sender, EventArgs e) {
			StartPinta(CM: true);
			bStartPintaCM.Enabled = false;
		}

		private void bStartPintaNormal_Click(object sender, EventArgs e) {
			StartPinta(CM: false);
			bStartPintaNormal.Enabled = false;
		}

		private void OpenStatusForm(string app, bool CM, Process process) {
			StatusForm sf = new StatusForm((int)numParticipant.Value, (int)numDay.Value, app, CM, process);
			sf.Show();
		}

		private void StartPinta(bool CM) {
			// Look for Pinta in the application folder + "/Pinta"
			string dir = Path.GetDirectoryName(Application.ExecutablePath);
			string path = Path.Combine(dir, CM ? "Pinta-CM" : "Pinta-Normal", "Pinta.exe");
			if (File.Exists(path)) {
				Process p = Process.Start(path);
				OpenStatusForm("Pinta", CM, p);
			} else {
				MessageBox.Show(string.Format("Error: Couldn't find Pinta in location {0}", path),
					"Application missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private enum INSTALLSTATE {
			INSTALLSTATE_NOTUSED = -7,  // component disabled
			INSTALLSTATE_BADCONFIG = -6,  // configuration data corrupt
			INSTALLSTATE_INCOMPLETE = -5,  // installation suspended or in progress
			INSTALLSTATE_SOURCEABSENT = -4,  // run from source, source is unavailable
			INSTALLSTATE_MOREDATA = -3,  // return buffer overflow
			INSTALLSTATE_INVALIDARG = -2,  // invalid function argument
			INSTALLSTATE_UNKNOWN = -1,  // unrecognized product or feature
			INSTALLSTATE_BROKEN = 0,  // broken
			INSTALLSTATE_ADVERTISED = 1,  // advertised feature
			INSTALLSTATE_REMOVED = 1,  // component being removed (action state, not settable)
			INSTALLSTATE_ABSENT = 2,  // uninstalled (or action state absent but clients remain)
			INSTALLSTATE_LOCAL = 3,  // installed on local drive
			INSTALLSTATE_SOURCE = 4,  // run from source, CD or net
			INSTALLSTATE_DEFAULT = 5,  // use default, local or source
		}

		[DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private extern static INSTALLSTATE MsiLocateComponent(string component, StringBuilder path, ref uint pathSize);

		private void StartWord2007(bool CM) {
			SetCommandMapEnabled(CM);

			uint size = 300;
			StringBuilder sb = new StringBuilder((int)size);
			var installstate = MsiLocateComponent("{0638C49D-BB8B-4CD1-B191-051E8F325736}", sb, ref size);
			if (installstate == INSTALLSTATE.INSTALLSTATE_LOCAL) {
				Process.Start(sb.ToString());
				OpenStatusForm("Word", CM);
			} else {
				MessageBox.Show("Error: Word 2007 not installed!", "Application missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public static void SetCommandMapEnabled(bool value) {
			RegistryKey key = Registry.CurrentUser.CreateSubKey("WordCommandMap");
			key.SetValue("CMEnabled", value ? 1 : 0, RegistryValueKind.DWord);
			key.Close();
		}
	}
}
