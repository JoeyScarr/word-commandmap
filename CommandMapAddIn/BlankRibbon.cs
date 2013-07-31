﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new BlankRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace CommandMapAddIn {
	[ComVisible(true)]
	public class BlankRibbon : Office.IRibbonExtensibility {
		private Office.IRibbonUI ribbon;
		private Word.Application m_Word;

		public BlankRibbon() {
		}

		public Word.Application Application {
			get { return m_Word; }
			set { m_Word = value; }
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID) {
			return GetResourceText("CommandMapAddIn.BlankRibbon.xml");
		}

		#endregion

		#region Ribbon Callbacks
		//Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

		public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
			this.ribbon = ribbonUI;
		}

		public void DisableClicked(Office.IRibbonControl control) {
			GlobalSettings.SetCommandMapEnabled(false);
			m_Word.Quit();
		}

		#endregion

		#region Helpers

		private static string GetResourceText(string resourceName) {
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i) {
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
						if (resourceReader != null) {
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
