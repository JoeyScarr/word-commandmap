using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CommandMapAddIn {
	public class ActivationButton : PerPixelAlphaForm {
		private WordInstance m_WordInstance;

		public ActivationButton()
			: base() {
			TopMost = true;
			ShowInTaskbar = false;
			FormBorderStyle = FormBorderStyle.None;
			Cursor = Cursors.Hand;

			// Add text to the bitmap
			Bitmap b = Properties.Resources.ActivationButton;
			Graphics g = Graphics.FromImage(b);
			g.DrawString("More commands... <Ctrl>", new Font("Segoe UI", 9F),
				new SolidBrush(Color.FromArgb(21, 66, 139)), new PointF(5, 1));
			SetBitmap(b);
		}

		public ActivationButton(WordInstance word)
			: this() {
			m_WordInstance = word;
		}

		public new void Show() {
			FollowWordPosition();
			base.Show();
		}

		protected override bool ShowWithoutActivation {
			get { return true; }
		}

		const int WS_EX_NOACTIVATE = 0x08000000;

		protected override CreateParams CreateParams {
			get {
				CreateParams param = base.CreateParams;
				param.ExStyle |= WS_EX_NOACTIVATE;
				return param;
			}
		}

		private void FollowWordPosition() {
			Rectangle windowRect = m_WordInstance.GetWindowPosition();
			Left = windowRect.Left + 200;
			Top = windowRect.Top + GlobalSettings.TITLEBAR_HEIGHT + GlobalSettings.BASE_RIBBON_HEIGHT;
		}
	}
}
