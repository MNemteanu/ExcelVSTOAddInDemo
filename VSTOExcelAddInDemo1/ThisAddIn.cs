using System;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

using VBA = Microsoft.Vbe.Interop;


namespace VSTOExcelAddInDemo1
{
	public partial class ThisAddIn
	{
		private static string _latestVBACode;
		private static int _latestVBACodeVersion;
		private static string _vbaCodeFile = "VBATest.txt";
		private static string _vbaCodeKey = "VBAVersion=";
		Excel._Workbook _activeWorkbook;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			this.Application.WorkbookActivate += Application_WorkbookActivate;
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		private void Application_WorkbookActivate(Excel.Workbook Wb)
		{
			// Trying to call from the XML Ribbon button instead
			_activeWorkbook = this.Application.ActiveWorkbook;
		}

		private void Application_WorkbookOpen(Excel.Workbook Wb)
		{
			throw new NotImplementedException();
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return new MyXMLRibbon();
		}

		#endregion

		// My public methods

		public static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		// My private methods

		public string CheckVBACodeVersion(/*Excel._Workbook pWorkbook*/)
		{
			_latestVBACodeVersion = GetLatestVBACodeVersion();

			int installedVBACodeVersion = GetInstalledVBACodeVersion(_latestVBACodeVersion, _activeWorkbook);

			if (installedVBACodeVersion < _latestVBACodeVersion)
			{
				UpdateOrInstallVBACode(_activeWorkbook);
			}

			return "checked"; // Return some information here like "VBA macro is up to date", "VBA macro version was updated to version ..."
		}

		private int GetInstalledVBACodeVersion(int pCurrentVersion, Excel._Workbook pWorkbook)
		{
			// Legacy method name: VBACodeVersion

			// Returns current installed VBA code version in the user's spreadsheet.
			// Removes old version if found

			int vbaCodeVersion = 0;

			var project = pWorkbook.VBProject;

			var projectName = project.Name;

			foreach (var component in project.VBComponents)
			{
				VBA.VBComponent vbComponent = (VBA.VBComponent)component;
				if (vbComponent != null)
				{
					string componentName = vbComponent.Name;
					var componentCode = vbComponent.CodeModule;
					int componentCodeLines = componentCode.CountOfLines;
					int line = 1;

					string oneLine = "";

					while (line < componentCodeLines)
					{

						// Looking for the version number in the format YYMMDDHHmm
						// This must be in the VBA code in Excel to detect if needs updated with the code in the ExcelVBA.txt that is included in this project.
						oneLine = componentCode.Lines[line, 1];

						// Only look for version if not already found
						if (vbaCodeVersion == 0)
						{
							int versionSearchResult = SearchVBACodeVersionInLine(oneLine);
							if (versionSearchResult > 0)
							{
								vbaCodeVersion = versionSearchResult;
							}
						}

						line++;
					}
				}
			}

			return vbaCodeVersion;
		}

		private int GetLatestVBACodeVersion()
		{
			int vbaCodeVersion = 0;

			try
			{
				StreamReader latestVBACodeStreamReader = GetLatestVBACodeStreamReader();

				string oneLine;
				while ((oneLine = latestVBACodeStreamReader.ReadLine()) != null)
				{
					int versionSearchResult = SearchVBACodeVersionInLine(oneLine);
					if (versionSearchResult > 0)
					{
						vbaCodeVersion = versionSearchResult;
						break;
					}
				}
				latestVBACodeStreamReader.Close();
			}
			catch (Exception ex)
			{
				string errorMessage = ex.Message;

				System.Windows.Forms.MessageBox.Show($@"Cannot get latest VBA code version.
Error: {errorMessage}");
			}

			return vbaCodeVersion;
		}

		private static StreamReader GetLatestVBACodeStreamReader()
		{
			_latestVBACode = GetResourceText($"VSTOExcelAddInDemo1.{_vbaCodeFile}");
			byte[] vbaCodeByteArray = Encoding.ASCII.GetBytes(_latestVBACode);
			MemoryStream vbaCodeMemoryStream = new MemoryStream(vbaCodeByteArray);
			StreamReader streamReader = new StreamReader(vbaCodeMemoryStream);

			return streamReader;
		}

		private void UpdateOrInstallVBACode(Excel._Workbook pWorkbook)
		{
			// Need to add code to remove the old VAB code if updating an older version

			try
			{
				var newStandardModule = pWorkbook.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);

				var newCodeModule = newStandardModule.CodeModule;
				int lineNum = newCodeModule.CountOfLines + 1;

				if (string.IsNullOrEmpty(_latestVBACode))
				{
					// The latestVBACode should be already assigned since it was read to find the installed version
					// so normally should not get here
					_latestVBACode = GetResourceText($"VSTOExcelAddInDemo1.{_vbaCodeFile}");
				}
				//+
				newCodeModule.InsertLines(lineNum, _latestVBACode);
			}
			catch (Exception ex)
			{
				string errorMessage = ex.Message;

				System.Windows.Forms.MessageBox.Show($@"Cannot install VBA code.
Error: {errorMessage}");
			}
		}

		private int SearchVBACodeVersionInLine(string pLine)
		{
			int vbaVersion = 0;

			if (pLine.ToUpper().Replace(" ", "").IndexOf(_vbaCodeKey.ToUpper()) >= 0)
			{
				string vbaVersionString = pLine.Substring(pLine.IndexOf('"') + 1, 10);
				Int32.TryParse(vbaVersionString, out vbaVersion);
			}

			return vbaVersion;
		}
	}
}
