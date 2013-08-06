using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CommandMapAddIn {
	public class ActivationButton : PerPixelAlphaForm {
		private WordInstance m_WordInstance;
		private Timer m_UpdateTimer;

		public ActivationButton()
			: base() {
			ShowInTaskbar = false;
			FormBorderStyle = FormBorderStyle.None;
			Cursor = Cursors.Hand;

			Enter += ActivationButton_Enter;
			Click += ActivationButton_Click;

			m_UpdateTimer = new Timer();
			m_UpdateTimer.Interval = 500; // Move every 500 ms
			m_UpdateTimer.Tick += m_UpdateTimer_Tick;
			m_UpdateTimer.Start();

			// Add text to the bitmap
			Bitmap b = Properties.Resources.ActivationTab;
			Graphics g = Graphics.FromImage(b);
			StringFormat format = new StringFormat();
			format.Alignment = StringAlignment.Center;
			g.DrawString("More commands... <Ctrl>", new Font("Segoe UI", 9F),
				new SolidBrush(Color.FromArgb(21, 66, 139)), new PointF(b.Width / 2, 1), format);
			SetBitmap(b);
		}

		void m_UpdateTimer_Tick(object sender, EventArgs e) {
			FollowWordPosition();
		}

		public ActivationButton(WordInstance word)
			: this() {
			m_WordInstance = word;
			WindowsApi.SetWindowParent(Handle, word.WindowHandle);
		}

		private void FocusWord() {
			m_WordInstance.Application.Activate();
		}

		void ActivationButton_Click(object sender, EventArgs e) {
			FocusWord();
		}

		void ActivationButton_Enter(object sender, EventArgs e) {
			FocusWord();
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
			Left = windowRect.Left + 25;
			Top = windowRect.Top + GlobalSettings.TITLEBAR_HEIGHT + GlobalSettings.BASE_RIBBON_HEIGHT + 1;
		}
	}
}
