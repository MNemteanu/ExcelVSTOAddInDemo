using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyXMLRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace VSTOExcelAddInDemo1
{
	[ComVisible(true)]
	public class MyXMLRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public MyXMLRibbon()
		{
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			//return GetResourceText("VSTOExcelAddInDemo1.MyXMLRibbon.xml");
			return VSTOExcelAddInDemo1.ThisAddIn.GetResourceText("VSTOExcelAddInDemo1.MyXMLRibbon.xml");
		}

		#endregion

		#region Ribbon Callbacks
		//Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		// Note for my callback methods:
		// A method like this can be called from more than on control (if for some reason want to do that)
		// and check the pControl.Id to decide what to do depending on what button was pressed.
		// Notice that if the method signature does not match the expected signature then the method won't be called (no error will be thrown).

		public void InstallVBA(Office.IRibbonControl pControl)
		{
			Globals.ThisAddIn.CheckVBACodeVersion();
		}

		public void TestVBACall_AddInMethod(Office.IRibbonControl pControl)
		{
			string workbookName = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
			Globals.ThisAddIn.Application.Workbooks[workbookName].Application.Run("TestVBACall_VBASub");
		}

		#endregion
	}
}
