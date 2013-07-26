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
			bool largeImageSet = false;
			if (icon != null) {
				if (icon.Width == 16) {
					if (item is RibbonButton) {
						((RibbonButton)item).SmallImage = icon;
					}
				} else {
					item.Image = icon;
					largeImageSet = true;
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
						largeImageSet = true;
					}
				}
			}
			// If there's no large image, use a small image if it exists
			if (!largeImageSet && icon != null) {
				item.Image = icon;
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

		private RibbonUpDown AddUpDown(RibbonItemCollection collection, string label, string msoImageName,
				decimal value, string suffix, decimal increment, decimal min = 0, decimal max = 100, EventHandler changed = null) {
			RibbonUpDown updown = new RibbonUpDown();
			updown.Text = label;
			updown.LabelWidth = 50;
			AssignImage(updown, msoImageName);
			decimal currentValue = value;
			updown.TextBoxText = value + suffix;
			updown.AllowTextEdit = true;
			updown.Enabled = true;
			updown.UpButtonClicked += new MouseEventHandler(delegate(object sender, MouseEventArgs e) {
				currentValue = Math.Min(currentValue + increment, max);
				updown.TextBoxText = currentValue + suffix;
			});
			updown.DownButtonClicked += new MouseEventHandler(delegate(object sender, MouseEventArgs e) {
				currentValue = Math.Max(currentValue - increment, min);
				updown.TextBoxText = currentValue + suffix;
			});
			updown.TextBoxTextChanged += changed;
			collection.Add(updown);
			return updown;
		}

		private RibbonLabel AddLabel(RibbonItemCollection collection, string text) {
			RibbonLabel label = new RibbonLabel();
			label.Text = text;
			collection.Add(label);
			return label;
		}

		private RibbonSeparator AddSeparator(RibbonItemCollection collection) {
			RibbonSeparator separator = new RibbonSeparator();
			collection.Add(separator);
			return separator;
		}

		private float GetLeadingFloat(string input) {
			return float.Parse(new string(input.Trim().TakeWhile(c => char.IsDigit(c) || c == '.').ToArray()));
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
			// TODO: Update everything in this panel when it changes
			AddLabel(panelParagraph.Items, "Indent");
			AddUpDown(panelParagraph.Items, "Left:", "IndentClassic",
				(decimal)m_WordInstance.Application.ActiveDocument.Paragraphs.LeftIndent, " cm", 0.1m, -27.9m, 55.8m,
				new EventHandler(delegate(object sender, EventArgs ea) {
					m_WordInstance.Application.ActiveDocument.Paragraphs.LeftIndent = GetLeadingFloat(((RibbonUpDown)sender).TextBoxText) * 28.35f; // cm to pt
				}));
			AddUpDown(panelParagraph.Items, "Right:", "ParagraphIndentRight",
				(decimal)m_WordInstance.Application.ActiveDocument.Paragraphs.RightIndent, " cm", 0.1m, -27.9m, 55.8m,
				new EventHandler(delegate(object sender, EventArgs ea) {
					m_WordInstance.Application.ActiveDocument.Paragraphs.RightIndent = GetLeadingFloat(((RibbonUpDown)sender).TextBoxText) * 28.35f; // cm to pt
				}));
			AddSeparator(panelParagraph.Items);
			AddLabel(panelParagraph.Items, "Spacing");
			AddUpDown(panelParagraph.Items, "Before:", "ParagraphSpacingIncrease",
				(decimal)m_WordInstance.Application.ActiveDocument.Paragraphs.SpaceBefore, " pt", 6, 0, 1584,
				new EventHandler(delegate(object sender, EventArgs ea) {
					m_WordInstance.Application.ActiveDocument.Paragraphs.SpaceBefore = GetLeadingFloat(((RibbonUpDown)sender).TextBoxText);
				}));
			AddUpDown(panelParagraph.Items, "After:", "ParagraphSpacingDecrease",
				(decimal)m_WordInstance.Application.ActiveDocument.Paragraphs.SpaceAfter, " pt", 6, 0, 1584,
				new EventHandler(delegate(object sender, EventArgs ea) {
					m_WordInstance.Application.ActiveDocument.Paragraphs.SpaceAfter = GetLeadingFloat(((RibbonUpDown)sender).TextBoxText);
				}));

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


			/*********************************************
			 * REVIEW TAB
			 *********************************************/
			// Proofing panel
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Spelling & Grammar", "Spelling", "Spelling");
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Research", "LookUp", "ResearchPane");
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Thesaurus", "Thesaurus", "Thesaurus");
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Translate", "Translate", "Translate");
			AddButton(panelProofing.Items, RibbonButtonStyle.DropDown, "Translation ScreenTip", "TranslationToolTip", "TranslationToolTip", RibbonElementSizeMode.Medium);
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Set Language", "SetLanguage", "SetLanguage", RibbonElementSizeMode.Medium);
			AddButton(panelProofing.Items, RibbonButtonStyle.Normal, "Word Count", "WordCountList", "WordCount", RibbonElementSizeMode.Medium);

			// Comments panel
			AddButton(panelComments.Items, RibbonButtonStyle.Normal, "New Comment", "ReviewNewComment", "ReviewNewComment");
			AddButton(panelComments.Items, RibbonButtonStyle.SplitDropDown, "Delete", "ReviewDeleteComment", "ReviewDeleteComment");
			AddButton(panelComments.Items, RibbonButtonStyle.Normal, "Previous", "ReviewPreviousComment", "ReviewPreviousComment");
			AddButton(panelComments.Items, RibbonButtonStyle.Normal, "Next", "ReviewNextComment", "ReviewNextComment");

			// Tracking panel
			AddButton(panelTracking.Items, RibbonButtonStyle.SplitDropDown, "Track Changes", "ReviewTrackChanges", "ReviewTrackChanges");
			AddButton(panelTracking.Items, RibbonButtonStyle.DropDown, "Balloons", "ReviewBalloonsMenu", "ReviewBalloonsMenu");
			// TODO: Add combo box here
			AddButton(panelTracking.Items, RibbonButtonStyle.DropDown, "Show Markup", "ReviewShowMarkupMenu", "ReviewShowMarkupMenu", RibbonElementSizeMode.Medium);
			AddButton(panelTracking.Items, RibbonButtonStyle.SplitDropDown, "Reviewing Pane", "ReviewReviewingPaneVertical", "ReviewReviewingPane", RibbonElementSizeMode.Medium);

			// Changes panel
			AddButton(panelChanges.Items, RibbonButtonStyle.SplitDropDown, "Accept", "ReviewAcceptChange", "ReviewAcceptChange");
			AddButton(panelChanges.Items, RibbonButtonStyle.SplitDropDown, "Reject", "ReviewRejectChange", "ReviewRejectChange");
			AddButton(panelChanges.Items, RibbonButtonStyle.Normal, "Previous", "ReviewPreviousChange", "ReviewPreviousChange", RibbonElementSizeMode.Medium);
			AddButton(panelChanges.Items, RibbonButtonStyle.Normal, "Next", "ReviewNextChange", "ReviewNextChange", RibbonElementSizeMode.Medium);

			// Compare panel
			AddButton(panelCompare.Items, RibbonButtonStyle.DropDown, "Compare", "ReviewCompareMenu", "ReviewCompareMenu");
			AddButton(panelCompare.Items, RibbonButtonStyle.DropDown, "Show Source Documents", "ReviewViewChangesInTheSourceDocument", "ReviewShowSourceDocumentsMenu"); // Wrong image, should be ReviewShowSourceDocumentsMenu

			// Protect panel
			AddButton(panelProtect.Items, RibbonButtonStyle.DropDown, "Protect Document", "ProtectDocument", "ReviewProtectDocumentMenu");


			/*********************************************
			 * VIEW TAB
			 *********************************************/
			// Document Views panel
			// TODO: Figure out how to do radio buttons

			// Show/Hide panel
			// TODO: Figure out how to do checkboxes

			// Zoom panel
			AddButton(panelZoom.Items, RibbonButtonStyle.Normal, "Zoom", "ZoomPrintPreviewExcel", "ZoomDialog");
			AddButton(panelZoom.Items, RibbonButtonStyle.Normal, "100%", "ZoomCurrent100", "Zoom100");
			AddButton(panelZoom.Items, RibbonButtonStyle.Normal, "One Page", "ZoomOnePage", "ZoomOnePage", RibbonElementSizeMode.Medium);
			AddButton(panelZoom.Items, RibbonButtonStyle.Normal, "Two Pages", "ZoomOnePage", "ZoomTwoPages", RibbonElementSizeMode.Medium); // Wrong image, should be ZoomTwoPages
			AddButton(panelZoom.Items, RibbonButtonStyle.Normal, "Page Width", "ZoomPageWidth", "ZoomPageWidth", RibbonElementSizeMode.Medium);

			// Window panel
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "New Window", "WindowNew", "WindowNew");
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "Arrange All", "WindowsArrangeAll", "WindowsArrangeAll");
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "Split", "WindowSplit", "WindowSplit");
			AddSeparator(panelWindow.Items);
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "View Side by Side", "WindowSideBySide", "WindowSideBySide", RibbonElementSizeMode.Medium);
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "Synchronous Scrolling", "WindowSideBySideSynchronousScrolling", "WindowSideBySideSynchronousScrolling", RibbonElementSizeMode.Medium);
			AddButton(panelWindow.Items, RibbonButtonStyle.Normal, "Reset Window Position", "WindowResetPosition", "WindowResetPosition", RibbonElementSizeMode.Medium);
			AddSeparator(panelWindow.Items);
			AddButton(panelWindow.Items, RibbonButtonStyle.DropDown, "Switch Windows", "WindowSwitchWindowsMenuExcel", "WindowSwitchWindowsMenuWord");

			// Macros panel
			AddButton(panelMacros.Items, RibbonButtonStyle.SplitDropDown, "Macros", "PlayMacro", "PlayMacro");
		}

		private void CommandMapForm_Leave(object sender, EventArgs e) {
			Hide();
		}

		private void CommandMapForm_Enter(object sender, EventArgs e) {
			m_WordInstance.Application.Activate();
		}
	}
}
