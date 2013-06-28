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
	public partial class CommandMapForm : Form {

		private WordInstance m_WordInstance;

		public CommandMapForm() {
			InitializeComponent();
		}

		public CommandMapForm(WordInstance instance)
			: this() {
				m_WordInstance = instance;
				FollowWordPosition();
		}

		private void button1_Click(object sender, EventArgs e) {
			FollowWordPosition();
		}

		private void FollowWordPosition() {
			Rectangle windowRect = m_WordInstance.GetWindowPosition();
			Location = windowRect.Location;
			Size = windowRect.Size;
		}
	}
}
