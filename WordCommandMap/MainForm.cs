using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordCommandMap {
	public partial class MainForm : Form {
		private string m_WordFilename = null;

		public MainForm() {
			InitializeComponent();
		}

		private void bStart_Click(object sender, EventArgs e) {
			// Spawn Word.
			WordInstance word = new WordInstance(m_WordFilename);

			// Spawn the CommandMap form, and attach it to the Word window.
			CommandMapForm cm = new CommandMapForm(word);
			cm.Show();
		}
	}
}
