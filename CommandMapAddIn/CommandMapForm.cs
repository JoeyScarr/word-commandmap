using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CommandMapAddIn {
	public partial class CommandMapForm : Form {

		private WordInstance m_WordInstance;
		private const int TITLEBAR_HEIGHT = 55;
		private const int STATUSBAR_HEIGHT = 22;

		public CommandMapForm() {
			InitializeComponent();
		}

		public CommandMapForm(WordInstance instance)
			: this() {

			m_WordInstance = instance;
			FollowWordPosition();
			BuildRibbon();
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

		public new void Show() {
			FollowWordPosition();
			base.Show();
			FollowWordPosition();
		}

		public new void Hide() {
			base.Hide();
		}

		private void AssignAction(RibbonItem item, string msoName) {
			if (item is RibbonButton) {
				string smallIconPath = string.Format(@"ribbon-icons\small\{0}.png", msoName);
				if (File.Exists(smallIconPath)) {
					((RibbonButton)item).SmallImage = Image.FromFile(smallIconPath);
				}
			}
			string largeIconPath = string.Format(@"ribbon-icons\large\{0}.png",msoName);
			if (File.Exists(largeIconPath)) {
				item.Image = Image.FromFile(largeIconPath);
			}
			item.Enabled = m_WordInstance.Application.CommandBars.GetEnabledMso(msoName);
			EventHandler handler = new EventHandler(delegate(object sender, EventArgs ea) {
				m_WordInstance.SendCommand(msoName);
			});
			item.Click += handler;
			item.DoubleClick += handler;
		}

		private void FollowWordPosition() {
			Rectangle windowRect = m_WordInstance.GetWindowPosition();
			Left = windowRect.Left;
			Top = windowRect.Top + TITLEBAR_HEIGHT;
			Width = windowRect.Width;
			Height = windowRect.Height - TITLEBAR_HEIGHT - STATUSBAR_HEIGHT;
		}

		private void BuildRibbon() {
			// Paste button
			RibbonButton paste = new RibbonButton();
			paste.Style = RibbonButtonStyle.SplitDropDown;
			paste.Text = "Paste";
			paste.Image = Image.FromFile(@"ribbon-icons\large\Paste.png");
			paste.SmallImage = Image.FromFile(@"ribbon-icons\small\Paste.png");
			AssignAction(paste, "Paste");
			clipboardPanel.Items.Add(paste);

			// Paste menu
			RibbonButton paste2 = new RibbonButton();
			paste2.Text = "Paste";
			paste2.Style = RibbonButtonStyle.Normal;
			paste2.SmallImage = Image.FromFile(@"ribbon-icons\small\Paste.png");
			AssignAction(paste2, "Paste");
			paste.DropDownItems.Add(paste2);
			RibbonButton pasteSpecial = new RibbonButton();
			pasteSpecial.Text = "Paste Special...";
			pasteSpecial.Style = RibbonButtonStyle.Normal;
			pasteSpecial.SmallImage = Image.FromFile(@"ribbon-icons\small\PasteSpecialDialog.png");
			AssignAction(pasteSpecial, "PasteSpecialDialog");
			paste.DropDownItems.Add(pasteSpecial);
			RibbonButton pasteAsHyperlink = new RibbonButton();
			pasteAsHyperlink.Text = "Paste as Hyperlink";
			pasteAsHyperlink.Style = RibbonButtonStyle.Normal;
			pasteAsHyperlink.SmallImage = Image.FromFile(@"ribbon-icons\small\PasteAsHyperlink.png");
			AssignAction(pasteAsHyperlink, "PasteAsHyperlink");
			paste.DropDownItems.Add(pasteAsHyperlink);

			// Cut button
			RibbonButton cut = new RibbonButton();
			cut.MaxSizeMode = RibbonElementSizeMode.Medium;
			cut.Style = RibbonButtonStyle.Normal;
			cut.Text = "Cut";
			cut.SmallImage = Image.FromFile(@"ribbon-icons\small\Cut.png");
			AssignAction(cut, "Cut");
			clipboardPanel.Items.Add(cut);

			// Copy button
			RibbonButton copy = new RibbonButton();
			copy.MaxSizeMode = RibbonElementSizeMode.Medium;
			copy.Style = RibbonButtonStyle.Normal;
			copy.Text = "Copy";
			copy.SmallImage = Image.FromFile(@"ribbon-icons\small\Copy.png");
			AssignAction(copy, "Copy");
			clipboardPanel.Items.Add(copy);

			// Paste button
			RibbonButton formatPainter = new RibbonButton();
			formatPainter.MaxSizeMode = RibbonElementSizeMode.Medium;
			formatPainter.Style = RibbonButtonStyle.Normal;
			formatPainter.Text = "Format Painter";
			formatPainter.SmallImage = Image.FromFile(@"ribbon-icons\small\FormatPainter.png");
			AssignAction(formatPainter, "FormatPainter");
			clipboardPanel.Items.Add(formatPainter);
			
		}
	}
}
