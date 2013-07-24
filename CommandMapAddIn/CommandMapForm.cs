﻿using System;
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
		private const int BASE_RIBBON_HEIGHT = 93;
		private const int CM_RIBBON_HEIGHT = 118;
		private const int NUM_CM_RIBBONS = 6;
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
			Top = windowRect.Top + TITLEBAR_HEIGHT + BASE_RIBBON_HEIGHT;
			Width = windowRect.Width;
			Height = Math.Min(CM_RIBBON_HEIGHT * NUM_CM_RIBBONS, windowRect.Height - TITLEBAR_HEIGHT - STATUSBAR_HEIGHT - BASE_RIBBON_HEIGHT);
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
				Hide();
				m_WordInstance.Application.Activate();
				m_WordInstance.SendCommand(msoName);
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

		private RibbonSeparator AddSeparator(RibbonItemCollection collection) {
			RibbonSeparator separator = new RibbonSeparator();
			collection.Add(separator);
			return separator;
		}

		private void BuildRibbon() {
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

			// Text panel
			AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Text Box", "TextBoxInsert", "TextBoxInsertGallery");
			AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Quick Parts", "QuickPartsInsertGallery", "QuickPartsInsertGallery");
			AddButton(panelText.Items, RibbonButtonStyle.DropDown, "WordArt", "QuickStylesSets", "WordArtInsertGalleryClassic"); // Wrong image
			AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Drop Cap", "DropCapOptionsDialog", "DropCapInsertGallery");
			AddButton(panelText.Items, RibbonButtonStyle.SplitDropDown, "Signature Line", "SignatureLineInsert", "SignatureLineInsert", RibbonElementSizeMode.Medium);
			AddButton(panelText.Items, RibbonButtonStyle.Normal, "Date & Time", "DateAndTimeInsert", "DateAndTimeInsert", RibbonElementSizeMode.Medium);
			AddButton(panelText.Items, RibbonButtonStyle.SplitDropDown, "Object", "OleObjectctInsert", "OleObjectctInsert", RibbonElementSizeMode.Medium);

			// Symbols panel
			AddButton(panelSymbols.Items, RibbonButtonStyle.SplitDropDown, "Equation", "AutoSum", "EquationInsertNew"); // Wrong image
			AddButton(panelSymbols.Items, RibbonButtonStyle.DropDown, "Symbol", "SymbolInsert", "SymbolInsertGallery");


			/*********************************************
			 * PAGE LAYOUT TAB
			 *********************************************/
			// Themes panel
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Themes", "ThemesGallery", "ThemesGallery");
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Colors", "ThemeColorsGallery", "ThemeColorsGallery", RibbonElementSizeMode.Medium);
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Fonts", "ThemeFontsGallery", "ThemeFontsGallery", RibbonElementSizeMode.Medium);
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Effects", "ThemeEffectsGallery", "ThemeEffectsGallery", RibbonElementSizeMode.Medium);

			// Page Setup panel
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Margins", "PageMarginsGallery", "PageMarginsGallery");
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Orientation", "PageOrientationGallery", "PageOrientationGallery");
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Size", "PageSizeGallery", "PageSizeGallery");
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Columns", "ColumnsDialog", "TableColumnsGallery");
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Breaks", "PageBreakInsertOrRemove", "BreaksGallery", RibbonElementSizeMode.Medium);
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Line Numbers", "LineNumbersMenu", "LineNumbersMenu", RibbonElementSizeMode.Medium);
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Hyphenation", "HyphenationOptions", "HyphenationMenu", RibbonElementSizeMode.Medium);

			// Page Background panel
			AddButton(panelPageBackground.Items, RibbonButtonStyle.DropDown, "Watermark", "WatermarkGallery", "WatermarkGallery");
			AddButton(panelPageBackground.Items, RibbonButtonStyle.DropDown, "Page Color", "PageColorPicker", "PageColorPicker");
			AddButton(panelPageBackground.Items, RibbonButtonStyle.Normal, "Page Borders", "BordersShadingDialogWord", "PageBorderAndShadingDialog");

			// Paragraph panel
			// TODO: Figure out how to do counter widgets

			// Arrange panel
			AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Position", "PicturePositionGallery", "PicturePositionGallery");
			AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Bring to Front", "ObjectBringToFront", "ObjectBringToFront");
			AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Send to Back", "ObjectSendToBack", "ObjectSendToBack");
			AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Text Wrapping", "TextWrappingMenu", "TextWrappingMenu");
			AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Align", "ObjectAlignMenu", "ObjectAlignMenu");
			AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Group", "ObjectsGroup", "ObjectsGroup");
			AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Rotate", "ObjectRotateGallery", "ObjectRotateGallery");


			/*********************************************
			 * REFERENCES TAB
			 *********************************************/
			// Table of Contents panel
			AddButton(panelTableOfContents.Items, RibbonButtonStyle.DropDown, "Table of Contents", "TableOfContentsGallery", "TableOfContentsGallery");
			AddButton(panelTableOfContents.Items, RibbonButtonStyle.DropDown, "Add Text", "TableOfContentsUpdate", "TableOfContentsAddTextGallery", RibbonElementSizeMode.Medium); // Wrong image, should be TableOfContentsAddTextGallery
			AddButton(panelTableOfContents.Items, RibbonButtonStyle.Normal, "Update Table", "TableOfContentsUpdate", "TableOfContentsUpdate", RibbonElementSizeMode.Medium);

			// Footnotes panel
			AddButton(panelFootnotes.Items, RibbonButtonStyle.Normal, "Insert Footnote", "FootnoteInsert", "FootnoteInsert");
			AddButton(panelFootnotes.Items, RibbonButtonStyle.Normal, "Insert Endnote", "EndnoteInsertWord", "EndnoteInsertWord", RibbonElementSizeMode.Medium);
			AddButton(panelFootnotes.Items, RibbonButtonStyle.SplitDropDown, "Next Footnote", "FootnoteNextWord", "FootnoteNextWord", RibbonElementSizeMode.Medium);
			AddButton(panelFootnotes.Items, RibbonButtonStyle.Normal, "Show Notes", "FootnotesEndnotesShow", "FootnotesEndnotesShow", RibbonElementSizeMode.Medium);

			// Captions panel
			AddButton(panelCaptions.Items, RibbonButtonStyle.Normal, "Insert Caption", "CaptionInsert", "CaptionInsert");
			AddButton(panelCaptions.Items, RibbonButtonStyle.Normal, "Insert Table of Figures", "TableOfFiguresInsert", "TableOfFiguresInsert", RibbonElementSizeMode.Medium);
			AddButton(panelCaptions.Items, RibbonButtonStyle.Normal, "Update Table", "TableOfContentsUpdate", "TableOfFiguresUpdate", RibbonElementSizeMode.Medium);
			AddButton(panelCaptions.Items, RibbonButtonStyle.Normal, "Cross-reference", "CrossReferenceInsert", "CrossReferenceInsert", RibbonElementSizeMode.Medium);

			// Index panel
			AddButton(panelIndex.Items, RibbonButtonStyle.Normal, "Mark Entry", "IndexMarkEntry", "IndexMarkEntry");
			AddButton(panelIndex.Items, RibbonButtonStyle.Normal, "Insert Index", "IndexInsert", "IndexInsert", RibbonElementSizeMode.Medium);
			AddButton(panelIndex.Items, RibbonButtonStyle.Normal, "Update Index", "TableOfContentsUpdate", "IndexUpdate", RibbonElementSizeMode.Medium);

			// Table of Authorities panel
			AddButton(panelTableOfAuthorities.Items, RibbonButtonStyle.Normal, "Mark Citation", "CitationMark", "CitationMark");
			AddButton(panelTableOfAuthorities.Items, RibbonButtonStyle.Normal, "Insert Table of Authorities", "TableOfAuthoritiesInsert", "TableOfAuthoritiesInsert", RibbonElementSizeMode.Medium);
			AddButton(panelTableOfAuthorities.Items, RibbonButtonStyle.Normal, "Update Table", "TableOfContentsUpdate", "TableOfAuthoritiesUpdate", RibbonElementSizeMode.Medium);


			/*********************************************
			 * MAILINGS TAB
			 *********************************************/
			// Create panel
			AddButton(panelCreate.Items, RibbonButtonStyle.Normal, "Envelopes", "EnvelopesAndLabelsDialog", "EnvelopesAndLabelsDialog");
			AddButton(panelCreate.Items, RibbonButtonStyle.Normal, "Labels", "LabelsDialog", "LabelsDialog");

			// Start Mail Merge panel
			AddButton(panelStartMailMerge.Items, RibbonButtonStyle.DropDown, "Start Mail Merge", "MailMergeStartMailMergeMenu", "MailMergeStartMailMergeMenu");
			AddButton(panelStartMailMerge.Items, RibbonButtonStyle.DropDown, "Select Recipients", "MailMergeSelectRecipients", "MailMergeSelectRecipients");
			AddButton(panelStartMailMerge.Items, RibbonButtonStyle.Normal, "Edit Recipient List", "MailMergeRecipientsEditList", "MailMergeRecipientsEditList");

			// Write & Insert Fields panel
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.Normal, "Highlight Merge Fields", "MailMergeHighlightMergeFields", "MailMergeHighlightMergeFields");
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.Normal, "Address Block", "MailMergeAddressBlockInsert", "MailMergeAddressBlockInsert");
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.Normal, "Greeting Line", "MailMergeGreetingLineInsert", "MailMergeGreetingLineInsert");
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.DropDown, "Insert Merge Field", "MailMergeMergeFieldInsert", "MailMergeMergeFieldInsert");
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.DropDown, "Rules", "MailMergeRules", "MailMergeRules", RibbonElementSizeMode.Medium);
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.Normal, "Match Fields", "MailMergeMatchFields", "MailMergeMatchFields", RibbonElementSizeMode.Medium);
			AddButton(panelWriteAndInsertFields.Items, RibbonButtonStyle.Normal, "Update Labels", "Refresh", "MailMergeUpdateLabels", RibbonElementSizeMode.Medium);

			// Preview Results panel
			AddButton(panelPreviewResults.Items, RibbonButtonStyle.Normal, "Preview Results", "MailMergeResultsPreview", "MailMergeResultsPreview");
			// TODO: Figure out how to do counter widgets
			AddSeparator(panelPreviewResults.Items);
			AddButton(panelPreviewResults.Items, RibbonButtonStyle.Normal, "Find Recipient", "Magnifier", "MailMergeFindRecipient", RibbonElementSizeMode.Medium); // Wrong image, should be MailMergeFindRecipient
			AddButton(panelPreviewResults.Items, RibbonButtonStyle.Normal, "Auto Check for Errors", "MailMergeAutoCheckForErrors", "MailMergeAutoCheckForErrors", RibbonElementSizeMode.Medium);

			// Finish panel
			AddButton(panelFinish.Items, RibbonButtonStyle.DropDown, "Finish & Merge", "MailMergeFinishAndMergeMenu", "MailMergeFinishAndMergeMenu");

		}

		private void CommandMapForm_Leave(object sender, EventArgs e) {
			Hide();
		}
	}
}
