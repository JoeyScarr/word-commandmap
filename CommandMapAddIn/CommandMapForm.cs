using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
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

		private void FollowWordPosition() {
			Rectangle windowRect = m_WordInstance.GetWindowPosition();
			Left = windowRect.Left;
			Top = windowRect.Top + TITLEBAR_HEIGHT;
			Width = windowRect.Width;
			Height = windowRect.Height - TITLEBAR_HEIGHT - STATUSBAR_HEIGHT;
		}

		private void AssignImage(RibbonItem item, string msoName) {
			// Get the first image
			Image icon = (Image)Properties.Resources.ResourceManager.GetObject(msoName);
			if (icon != null) {
				if (icon.Width == 16) {
					if (item is RibbonButton) {
						((RibbonButton)item).SmallImage = icon;
					}
				} else {
					item.Image = icon;
				}
			}
			// Get the second image
			Image icon1 = (Image)Properties.Resources.ResourceManager.GetObject(msoName + "1");
			if (icon1 != null) {
				if (icon1 != null) {
					if (icon1.Width == 16) {
						if (item is RibbonButton) {
							((RibbonButton)item).SmallImage = icon1;
						}
					} else {
						item.Image = icon1;
					}
				}
			}
		}

		private void AssignAction(RibbonItem item, string msoName) {
			item.Enabled = m_WordInstance.Application.CommandBars.GetEnabledMso(msoName);
			EventHandler handler = new EventHandler(delegate(object sender, EventArgs ea) {
				m_WordInstance.SendCommand(msoName);
				Hide();
			});
			item.Click += handler;
			item.DoubleClick += handler;
		}

		private RibbonButton AddButton(RibbonItemCollection collection, RibbonButtonStyle style, string label, string msoImageName,
			string msoCommandName, RibbonElementSizeMode maxSizeMode = RibbonElementSizeMode.None) {

			RibbonButton button = new RibbonButton();
			button.Style = style;
			button.MaxSizeMode = maxSizeMode;
			button.Text = label;
			AssignImage(button, msoImageName);
			AssignAction(button, msoCommandName);
			collection.Add(button);
			return button;
		}

		private void BuildRibbon() {
			/*********************************************
			 * HOME TAB
			 *********************************************/
			// Clipboard panel
			RibbonButton paste = AddButton(clipboardPanel.Items, RibbonButtonStyle.SplitDropDown, "Paste", "Paste", "Paste");
			AddButton(paste.DropDownItems, RibbonButtonStyle.Normal, "Paste", "Paste", "Paste");
			AddButton(paste.DropDownItems, RibbonButtonStyle.Normal, "Paste Special...", "PasteSpecialDialog", "PasteSpecialDialog");
			AddButton(paste.DropDownItems, RibbonButtonStyle.Normal, "Paste as Hyperlink", "PasteAsHyperlink", "PasteAsHyperlink");
			AddButton(clipboardPanel.Items, RibbonButtonStyle.Normal, "Cut", "Cut", "Cut", RibbonElementSizeMode.Medium);
			AddButton(clipboardPanel.Items, RibbonButtonStyle.Normal, "Copy", "Copy", "Copy", RibbonElementSizeMode.Medium);
			AddButton(clipboardPanel.Items, RibbonButtonStyle.Normal, "Format Painter", "FormatPainter", "FormatPainter", RibbonElementSizeMode.Medium);


			/*********************************************
			 * INSERT TAB
			 *********************************************/
			// Pages panel
			AddButton(panelPages.Items, RibbonButtonStyle.DropDown, "Cover Page", "CoverPageInsertGallery", "CoverPageInsertGallery");
			AddButton(panelPages.Items, RibbonButtonStyle.Normal, "Blank Page", "FileNew", "BlankPageInsert");
			AddButton(panelPages.Items, RibbonButtonStyle.Normal, "Page Break", "PageBreakInsertOrRemove", "PageBreakInsertWord");

			// Tables panel
			AddButton(panelTables.Items, RibbonButtonStyle.DropDown, "Table", "TableInsert", "TableInsertGallery");

			// Illustrations panel
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Picture", "PictureInsertFromFilePowerPoint", "PictureInsertFromFile");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Clip Art", "ClipArtInsert", "ClipArtInsert"); // Doesn't seem to work
			AddButton(panelIllustrations.Items, RibbonButtonStyle.DropDown, "Shapes", "ShapesMoreShapes", "GalleryAllShapesAndCanvas");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "SmartArt", "SmartArtInsert", "SmartArtInsert");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Chart", "ChartInsert", "ChartInsert");

			// Links panel
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Hyperlink", "HyperlinkInsert", "HyperlinkInsert");
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Bookmark", "FrontPageToggleBookmark", "BookmarkInsert");
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Cross-reference", "CrossReferenceInsert", "CrossReferenceInsert");

			// Header & Footer panel
			AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Header", "HeaderInsertGallery", "HeaderInsertGallery");
			AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Footer", "FooterInsertGallery", "FooterInsertGallery");
			AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Page Number", "PageNambersInFooterInsertGallery", "PageNambersInFooterInsertGallery");
			
		}
	}
}
