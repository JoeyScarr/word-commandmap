using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace CommandMapAddIn {
	public partial class ThisAddIn {

		private void ThisAddIn_Startup(object sender, System.EventArgs e) {
			// Create a WordInstance
			WordInstance word = new WordInstance(Application);

			// Spawn the CommandMap form, and attach it to the Word window.
			CommandMapForm cm = new CommandMapForm(word);
			cm.Show();
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
		}

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
			return new BlankRibbon();
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup() {
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
