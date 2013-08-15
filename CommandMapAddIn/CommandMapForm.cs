using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gma.UserActivityMonitor;
using Word = Microsoft.Office.Interop.Word;

namespace CommandMapAddIn {
	public partial class CommandMapForm : Form {

		private WordInstance m_WordInstance;
		private List<RibbonItem> m_Controls;

		public CommandMapForm() {
			InitializeComponent();

			m_Controls = new List<RibbonItem>();

			HookManager.MouseClick += HookManager_MouseClick;
		}

		public CommandMapForm(WordInstance instance)
			: this() {

			m_WordInstance = instance;

			FollowWordPosition();
			BuildRibbon();
		}

		private void HookManager_MouseClick(object sender, MouseEventArgs e) {
			Thread t = new Thread(new ThreadStart(delegate() {
				// Hide the CM if a click was detected outside the window.
				if (!Bounds.Contains(e.Location)) {
					Invoke(new System.Action(delegate() {
						Hide();
					}));
				}
			}));
			t.Start();
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
			foreach (RibbonItem item in m_Controls) {
				SetEnabled(item);
			}
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
			Top = windowRect.Top + GlobalSettings.TITLEBAR_HEIGHT + GlobalSettings.BASE_RIBBON_HEIGHT;
			Width = windowRect.Width;
			Height = Math.Min(GlobalSettings.CM_RIBBON_HEIGHT * GlobalSettings.NUM_CM_RIBBONS,
				windowRect.Height - GlobalSettings.TITLEBAR_HEIGHT - GlobalSettings.STATUSBAR_HEIGHT - GlobalSettings.BASE_RIBBON_HEIGHT);
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

		private void RunCommand(Action action) {
			Hide();
			Thread t = new Thread(new ThreadStart(action));
			t.Start();
			m_WordInstance.Focus();
		}

		private void RunCommand(string msoName) {
			RunCommand(delegate() {
				m_WordInstance.SendCommand(msoName);
			});
		}

		private void AssignAction(RibbonItem item, string msoName) {
			item.Tag = msoName;
			m_Controls.Add(item);
			EventHandler handler = new EventHandler(delegate(object sender, EventArgs ea) {
				RunCommand(msoName);
			});
			item.Click += handler;
			item.DoubleClick += handler;
			if (item is RibbonCheckBox) {
				((RibbonCheckBox)item).CheckBoxCheckChanged += handler;
			}
		}

		private void SetEnabled(RibbonItem item) {
			string msoName = (string)item.Tag;
			item.Enabled = msoName == "" ? false : m_WordInstance.Application.CommandBars.GetEnabledMso(msoName);
		}

		private RibbonButton AddButton(RibbonItemCollection collection, RibbonButtonStyle style, string label, string msoImageName,
			string msoCommandName, RibbonElementSizeMode maxSizeMode = RibbonElementSizeMode.None, Action customAction = null) {

			RibbonButton button = new RibbonButton();
			button.Style = style;
			button.MaxSizeMode = maxSizeMode;
			button.Text = label;
			AssignImage(button, msoImageName);
			if (customAction == null) {
				AssignAction(button, msoCommandName);
			} else {
				EventHandler handler = new EventHandler(delegate(object sender, EventArgs ea) {
					RunCommand(customAction);
				});
				button.Click += handler;
				button.DoubleClick += handler;
			}
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

		private RibbonCheckBox AddCheckBox(RibbonItemCollection collection, string text, string msoCommandName) {
			RibbonCheckBox checkbox = new RibbonCheckBox();
			checkbox.Text = text;
			collection.Add(checkbox);
			AssignAction(checkbox, msoCommandName);
			return checkbox;
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
			var coverPage = AddButton(panelPages.Items, RibbonButtonStyle.DropDown, "Cover Page", "CoverPageInsertGallery", "CoverPageInsertGallery");
			// MISSING: Cover Page gallery
			AddButton(coverPage.DropDownItems, RibbonButtonStyle.Normal, "Remove Current Cover Page", "CoverPageRemove", "CoverPageRemove");
			AddButton(coverPage.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Cover Page Gallery...", "SaveSelectionToCoverPageGallery", "SaveSelectionToCoverPageGallery");
			AddButton(panelPages.Items, RibbonButtonStyle.Normal, "Blank Page", "FileNew", "BlankPageInsert");
			AddButton(panelPages.Items, RibbonButtonStyle.Normal, "Page Break", "PageBreakInsertOrRemove", "PageBreakInsertWord");

			// Tables panel
			var table = AddButton(panelTables.Items, RibbonButtonStyle.DropDown, "Table", "TableInsert", "TableInsertGallery");
			// MISSING: Table drawing thingy
			AddButton(table.DropDownItems, RibbonButtonStyle.Normal, "Insert Table...", "TableInsert", "TableInsertDialogWord");
			AddButton(table.DropDownItems, RibbonButtonStyle.Normal, "Draw Table", "TableDrawTable", "TableDrawTable");
			AddButton(table.DropDownItems, RibbonButtonStyle.Normal, "Convert Text to Table...", "ConvertTextToTable", "ConvertTextToTable");
			AddButton(table.DropDownItems, RibbonButtonStyle.Normal, "Excel Spreadsheet", "TableExcelSpreadsheetInsert", "TableExcelSpreadsheetInsert");
			var quickTablesGallery = AddButton(table.DropDownItems, RibbonButtonStyle.DropDown, "Quick Tables", "TableInsert", "QuickTablesInsertGallery");
			// MISSING: Quick tables gallery
			AddButton(quickTablesGallery.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Quick Tables Gallery...", "SaveSelectionToQuickTablesGallery", "SaveSelectionToQuickTablesGallery");

			// Illustrations panel
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Picture", "PictureInsertFromFilePowerPoint", "PictureInsertFromFile");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Clip Art", "ClipArtInsert", "ClipArtInsert"); // Doesn't seem to work
			var shapes = AddButton(panelIllustrations.Items, RibbonButtonStyle.DropDown, "Shapes", "ShapesMoreShapes", "GalleryAllShapesAndCanvas");
			// MISSING: Shapes gallery
			AddButton(shapes.DropDownItems, RibbonButtonStyle.Normal, "New Drawing Canvas", "InsertDrawingCanvas", "DrawingCanvasInsert");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "SmartArt", "SmartArtInsert", "SmartArtInsert");
			AddButton(panelIllustrations.Items, RibbonButtonStyle.Normal, "Chart", "ChartInsert", "ChartInsert");

			// Links panel
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Hyperlink", "HyperlinkInsert", "HyperlinkInsert");
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Bookmark", "FrontPageToggleBookmark", "BookmarkInsert");
			AddButton(panelLinks.Items, RibbonButtonStyle.Normal, "Cross-reference", "CrossReferenceInsert", "CrossReferenceInsert");

			// Header & Footer panel
			var header = AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Header", "HeaderInsertGallery", "HeaderInsertGallery");
			// MISSING: Header gallery
			AddButton(header.DropDownItems, RibbonButtonStyle.Normal, "Edit Header", "HeaderInsertGallery", "HeaderFooterEditHeader");
			AddButton(header.DropDownItems, RibbonButtonStyle.Normal, "Remove Header", "HeaderFooterRemoveHeaderWord", "HeaderFooterRemoveHeaderWord");
			AddButton(header.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Header Gallery...", "SaveSelectionToHeaderGallery", "SaveSelectionToHeaderGallery");
			var footer = AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Footer", "FooterInsertGallery", "FooterInsertGallery");
			// MISSING: Footer gallery
			AddButton(footer.DropDownItems, RibbonButtonStyle.Normal, "Edit Footer", "FooterInsertGallery", "HeaderFooterEditFooter");
			AddButton(footer.DropDownItems, RibbonButtonStyle.Normal, "Remove Footer", "HeaderFooterRemoveFooterWord", "HeaderFooterRemoveFooterWord");
			AddButton(footer.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Footer Gallery...", "SaveSelectionToFooterGallery", "SaveSelectionToFooterGallery");
			var pageNumber = AddButton(panelHeaderFooter.Items, RibbonButtonStyle.DropDown, "Page Number", "PageNambersInFooterInsertGallery", "PageNambersInFooterInsertGallery");
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.DropDown, "Top of Page", "PageNumbersInHeaderInsertGallery", "PageNumbersInHeaderInsertGallery");
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.DropDown, "Bottom of Page", "PageNambersInFooterInsertGallery", "PageNambersInFooterInsertGallery");
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.DropDown, "Page Margins", "PageNambersInMarginsInsertGallery", "PageNambersInMarginsInsertGallery");
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.DropDown, "Current Position", "PageNambersInFooterInsertGallery", "PageNambersInFooterInsertGallery");
			// MISSING: Gallery submenus for all of the above
			AddSeparator(pageNumber.DropDownItems);
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.Normal, "Format Page Numbers...", "PageNumberFormat", "PageNumberFormat");
			AddButton(pageNumber.DropDownItems, RibbonButtonStyle.Normal, "Remove Page Numbers", "PageNumbersRemove", "PageNumbersRemove");

			// Text panel
			var textBox = AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Text Box", "TextBoxInsert", "TextBoxInsertGallery");
			// MISSING: Text box gallery
			AddButton(textBox.DropDownItems, RibbonButtonStyle.Normal, "Draw Text Box", "TextBoxInsert", "TextBoxDrawMenu");
			AddButton(textBox.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Text Box Gallery", "SaveSelectionToTextBoxGallery", "SaveSelectionToTextBoxGallery");
			var quickParts = AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Quick Parts", "QuickPartsInsertGallery", "QuickPartsInsertGallery");
			AddButton(quickParts.DropDownItems, RibbonButtonStyle.DropDown, "Document Property", "PropertyInsert", "PropertyInsert");
			// MISSING: Document Properties gallery
			AddButton(quickParts.DropDownItems, RibbonButtonStyle.Normal, "Field...", "FieldInsert", "FieldInsert");
			AddSeparator(quickParts.DropDownItems);
			AddButton(quickParts.DropDownItems, RibbonButtonStyle.Normal, "Building Blocks Organizer...", "Organizer", "BuildingBlocksOrganizer");
			AddButton(quickParts.DropDownItems, RibbonButtonStyle.Normal, "Get More on Office Online...", "QuickPartsInsertFromOnline", "QuickPartsInsertFromOnline"); // Missing image
			AddButton(quickParts.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Quick Part Gallery...", "SaveSelectionToQuickPartGallery", "SaveSelectionToQuickPartGallery");
			var wordArt = AddButton(panelText.Items, RibbonButtonStyle.DropDown, "WordArt", "QuickStylesSets", "WordArtInsertGalleryClassic"); // Wrong image
			// MISSING: WordArt gallery
			var dropCap = AddButton(panelText.Items, RibbonButtonStyle.DropDown, "Drop Cap", "DropCapOptionsDialog", "DropCapInsertGallery");
			AddSeparator(dropCap.DropDownItems);
			AddButton(dropCap.DropDownItems, RibbonButtonStyle.Normal, "Drop Cap Options...", "DropCapOptionsDialog", "DropCapOptionsDialog");
			var signatureLine = AddButton(panelText.Items, RibbonButtonStyle.SplitDropDown, "Signature Line", "SignatureLineInsert", "SignatureLineInsert", RibbonElementSizeMode.Medium);
			AddButton(signatureLine.DropDownItems, RibbonButtonStyle.Normal, "Microsoft Office Signature Line...", "", "SignatureLineInsert");
			AddSeparator(signatureLine.DropDownItems);
			AddButton(signatureLine.DropDownItems, RibbonButtonStyle.Normal, "Add Signature Services...", "", "SignatureServicesAdd");
			AddButton(panelText.Items, RibbonButtonStyle.Normal, "Date & Time", "DateAndTimeInsert", "DateAndTimeInsert", RibbonElementSizeMode.Medium);
			var objectMenu = AddButton(panelText.Items, RibbonButtonStyle.SplitDropDown, "Object", "OleObjectctInsert", "OleObjectctInsert", RibbonElementSizeMode.Medium);
			AddButton(objectMenu.DropDownItems, RibbonButtonStyle.Normal, "Object...", "OleObjectctInsert", "OleObjectctInsert");
			AddButton(objectMenu.DropDownItems, RibbonButtonStyle.Normal, "Text from File...", "TextFromFileInsert", "TextFromFileInsert");

			// Symbols panel
			var equation = AddButton(panelSymbols.Items, RibbonButtonStyle.SplitDropDown, "Equation", "AutoSum", "EquationInsertNew"); // Wrong image
			// MISSING: Equation gallery
			AddSeparator(equation.DropDownItems);
			AddButton(equation.DropDownItems, RibbonButtonStyle.Normal, "Insert New Equation", "EquationOptions", "EquationInsertNew");
			AddButton(equation.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Equation Gallery...", "AutoSum", "SaveSelectionToEquationGallery"); // Wrong image
			var symbol = AddButton(panelSymbols.Items, RibbonButtonStyle.DropDown, "Symbol", "SymbolInsert", "SymbolInsertGallery");
			// MISSING: Symbol gallery
			AddSeparator(symbol.DropDownItems);
			AddButton(symbol.DropDownItems, RibbonButtonStyle.Normal, "More Symbols...", "SymbolInsert", "SymbolsDialog");


			/*********************************************
			 * PAGE LAYOUT TAB
			 *********************************************/
			// Themes panel
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Themes", "ThemesGallery", "ThemesGallery");
			// MISSING: Themes gallery
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Colors", "ThemeColorsGallery", "ThemeColorsGallery", RibbonElementSizeMode.Medium);
			// MISSING: Colors gallery
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Fonts", "ThemeFontsGallery", "ThemeFontsGallery", RibbonElementSizeMode.Medium);
			// MISSING: Fonts gallery
			AddButton(panelThemes.Items, RibbonButtonStyle.DropDown, "Effects", "ThemeEffectsGallery", "ThemeEffectsGallery", RibbonElementSizeMode.Medium);
			// MISSING: Effects gallery

			// Page Setup panel
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Margins", "PageMarginsGallery", "PageMarginsGallery");
			var orientation = AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Orientation", "PageOrientationGallery", "PageOrientationGallery");
			AddButton(orientation.DropDownItems, RibbonButtonStyle.Normal, "Portrait", "PageOrientationPortraitLandscape", "PageOrientationPortraitLandscape");
			AddButton(orientation.DropDownItems, RibbonButtonStyle.Normal, "Landscape", "PageOrientationPortraitLandscape", "PageOrientationPortraitLandscape");
			var size = AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Size", "PageSizeGallery", "PageSizeGallery");
			// MISSING: Size gallery
			AddSeparator(size.DropDownItems);
			AddButton(size.DropDownItems, RibbonButtonStyle.Normal, "More Paper Sizes...", "", "PageSizeMorePaperSizesDialog");
			var columns = AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Columns", "ColumnsDialog", "TableColumnsGallery");
			// MISSING: Columns gallery
			AddButton(columns.DropDownItems, RibbonButtonStyle.Normal, "More Columns...", "ColumnsDialog", "ColumnsDialog");
			AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Breaks", "PageBreakInsertOrRemove", "BreaksGallery", RibbonElementSizeMode.Medium);
			var lineNumbers = AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Line Numbers", "LineNumbersMenu", "LineNumbersMenu", RibbonElementSizeMode.Medium);
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "None", "LineNumbersOff", "LineNumbersOff");
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "Continuous", "LineNumbersContinuous", "LineNumbersContinuous");
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "Restart Each Page", "LineNumbersResetPage", "LineNumbersResetPage");
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "Restart Each Section", "LineNumbersResetSection", "LineNumbersResetSection");
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "Suppress for Current Paragraph", "LineNumbersSuppress", "LineNumbersSuppress");
			AddSeparator(lineNumbers.DropDownItems);
			AddButton(lineNumbers.DropDownItems, RibbonButtonStyle.Normal, "Line Numbering Options...", "LineNumbersOptionsDialog", "LineNumbersOptionsDialog");
			var hyphenation = AddButton(panelPageSetup.Items, RibbonButtonStyle.DropDown, "Hyphenation", "HyphenationOptions", "HyphenationMenu", RibbonElementSizeMode.Medium);
			AddButton(hyphenation.DropDownItems, RibbonButtonStyle.Normal, "None", "HyphenationNone", "HyphenationNone");
			AddButton(hyphenation.DropDownItems, RibbonButtonStyle.Normal, "Automatic", "HyphenationAutomatic", "HyphenationAutomatic");
			AddButton(hyphenation.DropDownItems, RibbonButtonStyle.Normal, "Manual", "HyphenationManual", "HyphenationManual");
			AddSeparator(hyphenation.DropDownItems);
			AddButton(hyphenation.DropDownItems, RibbonButtonStyle.Normal, "Hyphenation Options...", "HyphenationOptions", "HyphenationOptions");

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
			var position = AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Position", "PicturePositionGallery", "PicturePositionGallery");
			// MISSING: Position gallery
			AddSeparator(position.DropDownItems);
			AddButton(position.DropDownItems, RibbonButtonStyle.Normal, "More Layout Options...", "LayoutOptionsDialog", "LayoutOptionsDialog");
			var bringToFront = AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Bring to Front", "ObjectBringToFront", "ObjectBringToFront");
			AddButton(bringToFront.DropDownItems, RibbonButtonStyle.Normal, "Bring to Front", "ObjectBringToFront", "ObjectBringToFront");
			AddButton(bringToFront.DropDownItems, RibbonButtonStyle.Normal, "Bring Forward", "ObjectBringForward", "ObjectBringForward");
			AddButton(bringToFront.DropDownItems, RibbonButtonStyle.Normal, "Bring in Front of Text", "ObjectBringInFrontOfText", "ObjectBringInFrontOfText"); // Missing icon
			var sendToBack = AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Send to Back", "ObjectSendToBack", "ObjectSendToBack");
			AddButton(sendToBack.DropDownItems, RibbonButtonStyle.Normal, "Send to Back", "ObjectSendToBack", "ObjectSendToBack");
			AddButton(sendToBack.DropDownItems, RibbonButtonStyle.Normal, "Send Backward", "ObjectSendBackward", "ObjectSendBackward");
			AddButton(sendToBack.DropDownItems, RibbonButtonStyle.Normal, "Send Behind Text", "ObjectSendBehindText", "ObjectSendBehindText"); // Missing icon
			var textWrapping = AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Text Wrapping", "TextWrappingMenu", "TextWrappingMenu");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "In Line with Text", "TextWrappingInLineWithText", "TextWrappingInLineWithText");
			AddSeparator(textWrapping.DropDownItems);
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Square", "TextWrappingSquare", "TextWrappingSquare");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Tight", "TextWrappingTight", "TextWrappingTight");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Behind Text", "TextWrappingBehindText", "TextWrappingBehindText");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "In Front of Text", "TextWrappingInFrontOfText", "TextWrappingInFrontOfText");
			AddSeparator(textWrapping.DropDownItems);
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Top and Bottom", "TextWrappingTopAndBottom", "TextWrappingTopAndBottom");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Through", "TextWrappingMenuClassic", "TextWrappingThrough");
			AddSeparator(textWrapping.DropDownItems);
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "Edit Wrap Points", "TextWrappingEditWrapPoints", "TextWrappingEditWrapPoints");
			AddButton(textWrapping.DropDownItems, RibbonButtonStyle.Normal, "More Layout Options...", "LayoutOptionsDialog", "LayoutOptionsDialog");
			var align = AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Align", "ObjectAlignMenu", "ObjectAlignMenu");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Left", "ObjectsAlignLeft", "ObjectsAlignLeftSmart");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Center", "ObjectsAlignCenterHorizontal", "ObjectsAlignCenterHorizontalSmart");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Right", "ObjectsAlignRight", "ObjectsAlignRightSmart");
			AddSeparator(align.DropDownItems);
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Top", "ObjectsAlignTop", "ObjectsAlignTopSmart");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Middle", "ObjectsAlignMiddleVertical", "ObjectsAlignMiddleVerticalSmart");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Bottom", "ObjectsAlignBottom", "ObjectsAlignBottomSmart");
			AddSeparator(align.DropDownItems);
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Distribute Horizontally", "AlignDistributeHorizontallyClassic", "AlignDistributeHorizontally");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Distribute Vertically", "AlignDistributeVerticallyClassic", "AlignDistributeVertically");
			AddSeparator(align.DropDownItems);
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align to Page", "", "ObjectsAlignRelativeToContainerSmart");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align to Margin", "", "ObjectsAlignRelativeToMargin");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Align Selected Objects", "", "ObjectsAlignSelectedSmart");
			AddSeparator(align.DropDownItems);
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "View Gridlines", "", "ViewGridlines");
			AddButton(align.DropDownItems, RibbonButtonStyle.Normal, "Grid Settings...", "GridSettings", "GridSettings");
			var group = AddButton(panelArrange.Items, RibbonButtonStyle.SplitDropDown, "Group", "ObjectsGroup", "ObjectsGroup");
			AddButton(group.DropDownItems, RibbonButtonStyle.Normal, "Group", "ObjectsGroup", "ObjectsGroup");
			AddButton(group.DropDownItems, RibbonButtonStyle.Normal, "Group", "ObjectsRegroup", "ObjectsRegroup");
			AddSeparator(group.DropDownItems);
			AddButton(group.DropDownItems, RibbonButtonStyle.Normal, "Group", "ObjectsUngroup", "ObjectsUngroup");
			var rotate = AddButton(panelArrange.Items, RibbonButtonStyle.DropDown, "Rotate", "ObjectRotateGallery", "ObjectRotateGallery");
			AddButton(rotate.DropDownItems, RibbonButtonStyle.Normal, "Rotate Right 90°", "ObjectRotateRight90", "", customAction: delegate() {
				m_WordInstance.Application.ActiveWindow.Selection.ShapeRange.IncrementRotation(90);
			});
			AddButton(rotate.DropDownItems, RibbonButtonStyle.Normal, "Rotate Left 90°", "ObjectRotateLeft90", "", customAction: delegate() {
				m_WordInstance.Application.ActiveWindow.Selection.ShapeRange.IncrementRotation(-90);
			});
			AddButton(rotate.DropDownItems, RibbonButtonStyle.Normal, "Flip Vertical", "ObjectFlipVertical", "", customAction: delegate() {
				m_WordInstance.Application.ActiveWindow.Selection.ShapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipVertical);
			});
			AddButton(rotate.DropDownItems, RibbonButtonStyle.Normal, "Flip Horizontal", "ObjectFlipHorizontal", "", customAction: delegate() {
				m_WordInstance.Application.ActiveWindow.Selection.ShapeRange.Flip(Microsoft.Office.Core.MsoFlipCmd.msoFlipHorizontal);
			});
			// Note: None of the above seem to work.
			AddSeparator(rotate.DropDownItems);
			AddButton(rotate.DropDownItems, RibbonButtonStyle.Normal, "More Rotation Options...", "", "ObjectRotationOptionsDialog");


			/*********************************************
			 * REFERENCES TAB
			 *********************************************/
			// Table of Contents panel
			var tableOfContents = AddButton(panelTableOfContents.Items, RibbonButtonStyle.DropDown, "Table of Contents", "TableOfContentsGallery", "TableOfContentsGallery");
			AddButton(tableOfContents.DropDownItems, RibbonButtonStyle.Normal, "Insert Table of Contents...", "TableOfContentsDialog", "TableOfContentsDialog");
			AddButton(tableOfContents.DropDownItems, RibbonButtonStyle.Normal, "Remove Table of Contents", "TableOfContentsRemove", "TableOfContentsRemove");
			AddButton(tableOfContents.DropDownItems, RibbonButtonStyle.Normal, "Save Selection to Table of Contents Gallery...", "SaveSelectionToTableOfContentsGallery", "SaveSelectionToTableOfContentsGallery");
			var addText = AddButton(panelTableOfContents.Items, RibbonButtonStyle.DropDown, "Add Text", "TableOfContentsUpdate", "TableOfContentsAddTextGallery", RibbonElementSizeMode.Medium); // Wrong image, should be TableOfContentsAddTextGallery
			AddButton(addText.DropDownItems, RibbonButtonStyle.Normal, "Do Not Show in Table of Contents", "", "");
			AddButton(addText.DropDownItems, RibbonButtonStyle.Normal, "Level 1", "", "");
			AddButton(addText.DropDownItems, RibbonButtonStyle.Normal, "Level 2", "", "");
			AddButton(addText.DropDownItems, RibbonButtonStyle.Normal, "Level 3", "", "");
			AddButton(panelTableOfContents.Items, RibbonButtonStyle.Normal, "Update Table", "TableOfContentsUpdate", "TableOfContentsUpdate", RibbonElementSizeMode.Medium);

			// Footnotes panel
			AddButton(panelFootnotes.Items, RibbonButtonStyle.Normal, "Insert Footnote", "FootnoteInsert", "FootnoteInsert");
			AddButton(panelFootnotes.Items, RibbonButtonStyle.Normal, "Insert Endnote", "EndnoteInsertWord", "EndnoteInsertWord", RibbonElementSizeMode.Medium);
			var nextFootnote = AddButton(panelFootnotes.Items, RibbonButtonStyle.SplitDropDown, "Next Footnote", "FootnoteNextWord", "FootnoteNextWord", RibbonElementSizeMode.Medium);
			AddButton(nextFootnote.DropDownItems, RibbonButtonStyle.Normal, "Next Footnote", "FootnoteNextWord", "FootnoteNextWord");
			AddButton(nextFootnote.DropDownItems, RibbonButtonStyle.Normal, "Previous Footnote", "", "FootnotePreviousWord");
			AddButton(nextFootnote.DropDownItems, RibbonButtonStyle.Normal, "Next Endnote", "", "EndnoteNextWord");
			AddButton(nextFootnote.DropDownItems, RibbonButtonStyle.Normal, "Previous Endnote", "", "EndnotePreviousWord");
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
			// TODO: Convert these to radio buttons (not that important)
			AddButton(panelDocumentViews.Items, RibbonButtonStyle.Normal, "Print Layout", "ViewPrintLayoutView", "ViewPrintLayoutView");
			AddButton(panelDocumentViews.Items, RibbonButtonStyle.Normal, "Full Screen Reading", "ReadingMode", "ViewFullScreenReadingView");
			AddButton(panelDocumentViews.Items, RibbonButtonStyle.Normal, "Web Layout", "ViewWebLayoutView", "ViewWebLayoutView");
			AddButton(panelDocumentViews.Items, RibbonButtonStyle.Normal, "Outline", "ViewOutlineView", "ViewOutlineView");
			AddButton(panelDocumentViews.Items, RibbonButtonStyle.Normal, "Draft", "ViewDraftView", "ViewDraftView");

			// Show/Hide panel
			AddCheckBox(panelShowHide.Items, "Ruler", "ViewRulerWord");
			AddCheckBox(panelShowHide.Items, "Gridlines", "ViewGridlines");
			AddCheckBox(panelShowHide.Items, "Message Bar", "ViewMessageBar");
			AddCheckBox(panelShowHide.Items, "Document Map", "ViewDocumentMap");
			AddCheckBox(panelShowHide.Items, "Thumbnails", "ViewThumbnails");

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
			m_WordInstance.Focus();
		}

		private void panelPageSetup_ButtonMoreClick(object sender, EventArgs e) {
			RunCommand("PageSetupDialog");
		}

		private void panelParagraph_ButtonMoreClick(object sender, EventArgs e) {
			RunCommand("ParagraphDialog");
		}

		private void panelFootnotes_ButtonMoreClick(object sender, EventArgs e) {
			RunCommand("FootnoteEndnoteDialog");
		}
	}
}
